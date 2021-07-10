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
            DoGetCert       loMain  = null;
            tvProfile       loProfile = null;

            try
            {
                // Make sure any previous parent instance is shutdown before proceeding
                // (but only after everything has been installed and configured first).
                if ( File.Exists(Process.GetCurrentProcess().MainModule.FileName + ".config") )
                    Env.bKillProcessParent(Process.GetCurrentProcess());

                loProfile = tvProfile.oGlobal(DoGetCert.oHandleUpdates(ref args));

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

-NoIISBindingScript=''

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

    Set this switch False to leave the profile file untouched after a command line
    has been passed to the EXE and merged with the profile. When true, everything
    but command line keys will be saved. When false, not even status information
    will be written to the profile file (ie. '{INI}').

-ScriptSSO= SEE PROFILE FOR DEFAULT VALUE

    This is the PowerShell script that updates SSO servers with new certificates.

-ScriptStage1= SEE PROFILE FOR DEFAULT VALUE

    There are multiple stages involved with the process of getting a certificate
    from the certificate provider network. Each stage has an associated PowerShell
    script snippet. The stages are represented in this profile by -ScriptStage1
    thru -ScriptStage7.

-ServiceNameLive='LetsEncrypt'

    This is the name mapped to the live production certificate network service URL.

-ServiceNameStaging='LetsEncrypt-Staging'

    This is the name mapped to the non-production (ie. 'staging') certificate
    network service URL.

-ServiceReportEverything=True

    By default, all activity logged on the client during non-interactive mode
    is uploaded to the SCS server. This can be very helpful during automation
    testing. Once testing is complete, set this switch False to report errors
    only.

    Note: 'non-interactive mode' means the -Auto switch is set (see above).
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

-SubmissionRetries=42

    Pending submissions to the certificate provider network will be retried until
    they succeed or fail, at most this many times. By default, the  process will
    retry for 7 minutes (-SubmissionRetries times -SubmissionWaitSecs, see below)
    for challenge status updates as well as certificate request status updates.

-SubmissionWaitSecs=10

    These are the seconds of wait time after the DNS website challenge has been
    submitted to the certificate network as well as after the certificate request
    has been submitted. This is the amount of time during which the request should
    transition from a 'pending' state to anything other than 'pending'.

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
                            .Replace("{SsoKey}", gsSsoThumbprintFilestKey)
                            .Replace("{{", "{")
                            .Replace("}}", "}")
                            );

                    string lsFetchName = null;

                    // Fetch simple setup.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="Setup Application Folder.exe")
                            , loProfile.sRelativeToProfilePathFile(lsFetchName));

                    // Fetch source code.
                    if ( loProfile.bValue("-FetchSource", false) )
                        tvFetchResource.ToDisk(lsFetchName=System.Windows.Application.ResourceAssembly.GetName().Name, lsFetchName + ".zip", null);

                    // Updates start here.
                    if ( loProfile.bFileJustCreated )
                    {
                        loProfile["-Update2021-06-27"] = true;
                        loProfile.Save();
                    }
                    else
                    {
                        if ( !loProfile.bValue("-Update2021-06-27", false) )
                        {
                            // Force use of the latest version of ACME-PS.
                            File.Delete(loProfile.sRelativeToProfilePathFile("ACME-PS.zip"));
                            if ( Directory.Exists(loProfile.sRelativeToProfilePathFile("ACME-PS")) )
                                Directory.Delete(loProfile.sRelativeToProfilePathFile("ACME-PS"), true);

                            // Remove workspace init scripts (ie. use new defaults).
                            loProfile.Remove("-ScriptStage1");
                            loProfile.Remove("-ScriptStage1s");

                            loProfile["-Update2021-06-27"] = true;
                            loProfile.Save();
                        }
                    }
                    // Updates end here.

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
                                DoGetCert.ReportErrors(lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
                            }
                            else
                            {
                                DoGetCert.ReportEverything(lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
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
                Env.LogIt(Env.sExceptionMessage(ex));

                try
                {
                    tvProfile   loMinProfile = Env.oMinProfile(loProfile);
                    byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                    string      lsHash = HashClass.sHashIt(loMinProfile);
                    string      lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
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
                if ( null != Env.oSetupCertificate )
                {
                    // Get setup cert's private key file.
                    string  lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, Env.oSetupCertificate);
                            if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) && Env.sCertName(Env.oSetupCertificate) == Env.sNewClientSetupCertName )
                            {
                                // Remove cert's private key file (sadly, the OS typically let's these things accumulate forever).
                                File.Delete(lsMachineKeyPathFile);
                            }
                }

                moProfile["-CertificateSetupDone"] = value;
                moProfile.Save();
            }
        }


        private static tvMessageBox goStartupWaitMsg    = null;
        private static string gsCertificateKeyPrefix    = "-Cert_";
        private static string gsCertificateHashKey      = "-Hash";
        private static string gsContentKey              = "-Content";
        private static string gsHostProcess             = "GoPcBackup.exe";
        private static string gsSsoThumbprintFilestKey  = "-SsoThumbprintFiles";


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

                    tvProfile loClientCertificates = loClientCertificatesWithHash.oOneKeyProfile(gsCertificateKeyPrefix + "*", true);

                    foreach (X509Certificate2 loCertificate in loStore.Certificates)
                    {
                        if ( loClientCertificates.ContainsKey(loCertificate.Thumbprint) )
                        {
                            // The certificate's already in the store, remove it from the download list.
                            loClientCertificates.Remove(loCertificate.Thumbprint);
                        }
                        else
                        {
                            // The certificate is not in the download list, remove it from the store.
                            loStore.Remove(loCertificate);

                            loCertsAddedRemoved.Add("-Removed", String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                    , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint));

                            Env.LogIt(String.Format("    \"{0}\" removed.", loCertificate.Subject));
                        }
                    }

                    foreach (DictionaryEntry loEntry in loClientCertificates)
                    {
                        // Add new certificate to the store.
                        X509Certificate2    loCertificate = new X509Certificate2();
                                            loCertificate.Import(Convert.FromBase64String(new tvProfile(loEntry.Value.ToString()).sValue("-PublicKey", "")));

                        loStore.Add(loCertificate);

                        loCertsAddedRemoved.Add("-Added", String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint));

                        Env.LogIt(String.Format("    \"{0}\" added.", loCertificate.Subject));
                    }

                    foreach (X509Certificate2 loCertificate in loStore.Certificates)
                    {
                        loCertsAddedRemoved.Add("-All", String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint));
                    }
                }
                catch (Exception ex)
                {
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

            // Do nothing if this is an SSO server (or we are not doing thrumprint updates on this server).
            if ( DoGetCert.oDomainProfile.bValue("-IsSsoDomain", false) || moProfile.bValue("-SkipSsoThumbprintUpdates", false) )
                return true;

            // Do nothing if no SSO domain is defined for the current domain.
            string  lsSsoDnsName = DoGetCert.oDomainProfile.sValue("-SsoDnsName", "");
                    if ( "" == lsSsoDnsName )
                        return true;

            Env.LogIt("");
            Env.LogIt(String.Format("Checking for SSO thumbprint change (\"{0}\") ...", lsSsoDnsName));
            
            // At this point it is an error if no SSO thumbprint exists for the current domain.
            string      lsSsoThumbprint = DoGetCert.oDomainProfile.sValue("-SsoThumbprint", "");
                        if ( "" == lsSsoThumbprint )
                        {
                            Env.LogIt("");
                            throw new Exception(String.Format("The SSO certificate thumbprint for the \"{0}\" domain has not yet been set!", lsSsoDnsName));
                        }
            string      lsSsoPreviousThumbprint = moProfile.sValue("-SsoThumbprint", "");
                        if ( !moProfile.ContainsKey(gsSsoThumbprintFilestKey) )
                        {
                            string  lsDefaultPhysicalPathFiles = Path.Combine(moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot"), "web.config");
                            Site[]  loSetupCertBoundSiteArray = DoGetCert.oSetupCertBoundSiteArray(new ServerManager());
                                    moProfile.Remove(gsSsoThumbprintFilestKey + "Note");
                                    if ( null == loSetupCertBoundSiteArray || 0 == loSetupCertBoundSiteArray.Length )
                                    {
                                        Env.LogIt(String.Format("{0} will be set to the IIS default since no sites could be found bound to the current certificate.", gsSsoThumbprintFilestKey));

                                        if ( null != loSetupCertBoundSiteArray )
                                        {
                                            moProfile.Add(gsSsoThumbprintFilestKey, Environment.ExpandEnvironmentVariables(lsDefaultPhysicalPathFiles));
                                            moProfile.Add(gsSsoThumbprintFilestKey + "Note", "No IIS websites could be found bound to the current certificate.");
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
                                            moProfile.Add(gsSsoThumbprintFilestKey
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
            tvProfile   loFileList = new tvProfile();
                        foreach (DictionaryEntry loEntry in moProfile.oOneKeyProfile(gsSsoThumbprintFilestKey))
                        {
                            // Change gsSsoThumbprintFilestKey to "-Files".
                            loFileList.Add("-Files", loEntry.Value);
                        }
            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);
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
                                , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="ReplaceText.exe"), moProfile.sRelativeToProfilePathFile(lsFetchName));
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="ReplaceText.exe.txt"), moProfile.sRelativeToProfilePathFile(lsFetchName));

            System.Windows.Forms.Application.DoEvents();

            Env.bPowerScriptError = false;
            Env.sPowerScriptOutput = null;

            Env.LogIt("");
            Env.LogIt(String.Format("Replacing SSO thumbprint (\"{0}\") ...", lsSsoDnsName));
            Env.LogIt(loProcess.StartInfo.Arguments);

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
                    Env.LogIt(moProfile.sValue("-ReplaceTextTimeoutMsg", "*** thumbprint replacement sub-process timed-out ***\r\n\r\n"));

                int liExitCode = -1;
                    try { liExitCode = loProcess.ExitCode; } catch {}

                if ( Env.bPowerScriptError || liExitCode != 0 || !loProcess.HasExited )
                {
                    Env.LogIt(Env.sExceptionMessage("The thumbprint replacement sub-process experienced a critical failure."));
                }
                else
                {
                    lbReplaceSsoThumbprint = true;

                    Env.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            Env.bKillProcess(loProcess);

            if ( lbReplaceSsoThumbprint )
            {
                File.Delete(moProfile.sRelativeToProfilePathFile("ReplaceText.exe"));
                File.Delete(moProfile.sRelativeToProfilePathFile("ReplaceText.exe.txt"));

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
            else
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
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
            string  lsLoadBalancerPfxPathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "");

            Env.LogIt("");
            Env.LogIt("Checking for load balancer ...");

            // Do nothing if a load balancer PFX file location is not defined.
            if ( "" == lsLoadBalancerPfxPathFile )
            {
                Env.LogIt("No load balancer found.");
                return true;
            }

            Env.LogIt(String.Format("Copying new certificate for use on the \"{0}\" load balancer ...", asCertName));

            X509Certificate2 loLbCertificate = new X509Certificate2(asCertPathFile, asCertificatePassword
                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);

            Directory.CreateDirectory(Path.GetDirectoryName(lsLoadBalancerPfxPathFile));
            File.WriteAllBytes(lsLoadBalancerPfxPathFile
                    , loLbCertificate.Export(X509ContentType.Pfx, HashClass.sDecrypted(Env.oSetupCertificate, DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPassword", ""))));

            Env.LogSuccess();

            string  lsLoadBalancerExePathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerExePathFile", "");
                    if ( "" == lsLoadBalancerExePathFile && !this.bMainLoopStopped )
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

            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = lsLoadBalancerExePathFile;
                        loProcess.StartInfo.Arguments = lsLoadBalancerPfxPathFile;
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

                if ( Env.bPowerScriptError || liExitCode != 0 || !loProcess.HasExited )
                {
                    Env.LogIt(Env.sExceptionMessage("The load balancer certificate replacement sub-process experienced a critical failure."));
                }
                else
                {
                    File.Delete(lsLoadBalancerPfxPathFile);

                    lbUploadCertToLoadBalancer = true;

                    Env.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            Env.bKillProcess(loProcess);

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

        public static void ClearCache()
        {
            // Clear the domain profile cache.
            goDomainProfile = null;
        }

        public void LoadBalancerReleaseCert()
        {
            if ( moProfile.bValue("-UseStandAloneMode", true) )
                return;

            Env.LogIt("");
            Env.LogIt("The load balancer admin asserts the new certificate is ready for general release during the next maintenance window ...");

            string  lsCaption = "Release Load Balancer Certificate";
            string  lsLoadBalancerPfxComputer = DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxComputer", "-LoadBalancerPfxComputer not defined");
            string  lsLoadBalancerPfxPathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "-LoadBalancerPfxPathFile not defined");
            bool    lbWrongServer = "" != lsLoadBalancerPfxComputer && lsLoadBalancerPfxComputer != Env.sComputerName;
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
                lbReleased = false;

                this.ShowError(String.Format("The load balancer certificate (\"{0}\") on \"{1}\" can't be removed! ({2})"
                                                , lsLoadBalancerPfxPathFile, Env.sComputerName, ex.Message), lsCaption);

                DoGetCert.ReportErrors(lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
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

                    this.Show("Load balancer \"Certificate Ready\" notifications have been sent.", lsCaption
                            , tvMessageBoxButtons.OK, tvMessageBoxIcons.Done, "LoadBalancerCertReleased");
                }
            }
        }

        public void LogStage(string asStageId)
        {
            Env.LogIt("");
            Env.LogIt("");
            Env.LogIt(String.Format("Stage {0} ...", asStageId));
            Env.LogIt("");
        }

        public static void ReportErrors(string asHash, byte[] abtArrayMinProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(Env.sLogPathFile);
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return;

            string  lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                    if ( File.Exists(lsPreviousErrorsLogPathFile) )
                        File.AppendAllText(lsPreviousErrorsLogPathFile, asLogFileTextReported);
                    else
                        File.Copy(lsLogPathFile, lsPreviousErrorsLogPathFile, false);

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            loCompressContentProfile.Add(gsContentKey, File.ReadAllText(lsPreviousErrorsLogPathFile));

                loGetCertServiceClient.ReportErrors(asHash, abtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                File.Delete(lsPreviousErrorsLogPathFile);
            }
        }

        public static void ReportEverything(string asHash, byte[] abtArrayMinProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(Env.sLogPathFile);
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) || !tvProfile.oGlobal().bValue("-ServiceReportEverything", true) )
                return;

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            loCompressContentProfile.Add(gsContentKey, asLogFileTextReported);

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
            Env.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowError(this.oUI, asMessageText, asMessageCaption);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }
     
        private string sSingleSessionScriptPathFile
        {
            get
            {
                if ( null == msSingleSessionScriptPathFile && moProfile.bValue("-SingleSessionEnabled", false) )
                {
                    msSingleSessionScriptPathFile = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-PowerScriptSessionPathFile", "InGetCertSession.ps1"));

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

        private static bool bDoCheckIn(string asHash, byte[] abtArrayMinProfile)
        {
            bool lbDoCheckIn = false;

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return true;

            Env.LogIt("");
            Env.LogIt("Attempting check-in with certificate repository ...");

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                if ( null != goStartupWaitMsg )
                    goStartupWaitMsg.Hide();

                if ( null != goStartupWaitMsg && !tvProfile.oGlobal().bExit )
                    goStartupWaitMsg.Show();

                string      lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
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
                }
            }

            return lbDoCheckIn;
        }

        private static bool bDoCheckOut(string asHash, byte[] abtArrayMinProfile)
        {
            bool lbDoCheckOut = false;

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return true;

            Env.LogIt("Checking-out of certificate repository ...");

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                lbDoCheckOut = loGetCertServiceClient.bClientCheckOut(asHash, abtArrayMinProfile);
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

            return lbDoCheckOut;
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

            return Env.bRunPowerScript(out lsOutput, gbMainLoopStopped, null, asScript, abOpenOrCloseSingleSession, false);
        }
        private bool bRunPowerScript(out string asOutput, string asScript)
        {
            return Env.bRunPowerScript(out asOutput, gbMainLoopStopped, this.sSingleSessionScriptPathFile, asScript, false, false);
        }

        private bool bApplyNewCertToSSO(X509Certificate2 aoNewCertificate, X509Certificate2 aoOldCertificate, string asHash, byte[] abtArrayMinProfile)
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
                        if ( !Env.bRunPowerScript(out lsOutput, gbMainLoopStopped, null, moProfile.sValue("-ScriptSsoVerChk1", "try {(Get-AdfsProperties).HostName} catch {}"), false, true) )
                        {
                            Env.LogIt("Failed checking SSO vs. proxy.");
                            return false;
                        }

                        lbIsProxy = String.IsNullOrEmpty(lsOutput);

                        if ( lbIsProxy )
                        {
                            if ( !Env.bRunPowerScript(out lsOutput, gbMainLoopStopped, null, moProfile.sValue("-ScriptSsoProxyVerChk", "try {(Get-WebApplicationProxySslCertificate).HostName} catch {}"), false, true) )
                            {
                                Env.LogIt("Failed checking SSO proxy version.");
                                return false;
                            }

                            lbIsOlder = String.IsNullOrEmpty(lsOutput);
                        }
                        else
                        {
                            if ( !Env.bRunPowerScript(out lsOutput, gbMainLoopStopped, null, moProfile.sValue("-ScriptSsoVerChk2", "try {(Get-AdfsProperties).AuditLevel} catch {}"), false, true) )
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

if ( (Get-AdfsCertificate -CertificateType ""Token-Signing"").Thumbprint -eq ""{NewCertificateThumbprint}"" )
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
            string  lsNewSsoThumbprint = DoGetCert.oDomainProfile.sValue("-SsoThumbprint", "");
                    if ( "" == lsNewSsoThumbprint )
                    {
                        Env.LogIt("");
                        throw new Exception(String.Format("The SSO certificate thumbprint for the \"{0}\" domain has not yet been set!", Env.sCurrentCertificateName));
                    }
            string  lsCurrentThumbprint = Env.oCurrentCertificate().Thumbprint;

            if ( lbIsProxy && lsNewSsoThumbprint == lsCurrentThumbprint )
            {
                // This means the certificate on the SSO server hasn't changed yet.
                // The certificate on the proxy can't be changed until that happens.

                DateTime ldtTimeout = DateTime.Now.AddMinutes(moProfile.iValue("-SsoProxyTimeoutMins", 30));

                while ( !this.bMainLoopStopped && DateTime.Now < ldtTimeout )
                {
                    System.Windows.Forms.Application.DoEvents();

                    if ( !this.bMainLoopStopped )
                    {
                        int liSsoProxySleepSecs = moProfile.iValue("-SsoProxySleepSecs", 60);
                            if ( liSsoProxySleepSecs < 1 )
                                liSsoProxySleepSecs = 60;

                        Thread.Sleep(1000 * liSsoProxySleepSecs);
                        System.Windows.Forms.Application.DoEvents();

                        DoGetCert.ClearCache();
                        lsNewSsoThumbprint = DoGetCert.oDomainProfile.sValue("-SsoThumbprint", "");

                        if ( lsNewSsoThumbprint != lsCurrentThumbprint )
                            break;
                    }
                }

                if ( lsNewSsoThumbprint == lsCurrentThumbprint )
                {
                    Env.LogIt("");
                    throw new Exception(String.Format("Timeout waiting for a certificate change on the internal \"{0}\" server(s)!", Env.sCurrentCertificateName));
                }
            }

            lbApplyNewCertToSSO = this.bRunPowerScript(moProfile.sValue("-ScriptSSO", lsDefaultScript)
                    .Replace("{NewCertificateThumbprint}", aoNewCertificate.Thumbprint)
                    .Replace("{OldCertificateThumbprint}", aoOldCertificate.Thumbprint)
                    );

            if ( String.IsNullOrEmpty(lsDefaultScript) )
                throw new Exception("-ScriptSSO is empty. This is an error.");

            if ( lbApplyNewCertToSSO && !lbIsProxy )
            {
                Env.LogIt("");
                Env.LogIt("Updating SSO token signing certificate thumbprint in the database ...");

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
                    Env.LogIt(Env.sExceptionMessage("The SSO thumbprint update sub-process experienced a critical failure."));
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
                string  lsCertOverridePfxComputer = DoGetCert.oDomainProfile.sValue("-CertOverridePfxComputer", "");
                string  lsCertOverridePfxPathFile = DoGetCert.oDomainProfile.sValue("-CertOverridePfxPathFile", "");

                if ( "" != lsCertOverridePfxComputer + lsCertOverridePfxPathFile && ("" == lsCertOverridePfxComputer || "" == lsCertOverridePfxPathFile) )
                {
                    throw new Exception("Both CertOverridePfxComputer AND CertOverridePfxPathFile must be defined for certificate overrides to work properly.");
                }
                else
                if ( !DoGetCert.oDomainProfile.bValue("-CertOverridePfxReady", false) && Env.sComputerName == lsCertOverridePfxComputer && File.Exists(lsCertOverridePfxPathFile) )
                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
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
                        Env.LogIt(String.Format("The current certificate ({0}) appears to be invalid.", aoOldCertificate.Subject));
                        Env.LogIt("  (If you know otherwise, make sure internet connectivity is available and the system clock is accurate.)");
                        Env.LogIt("Expiration status will therefore be ignored and the process will now run.");
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
                                    int liRenewalDaysBeforeExpiration = DoGetCert.oDomainProfile.iValue("-RenewalDaysBeforeExpiration", 30);
                                        if ( moProfile.iValue("-RenewalDaysBeforeExpiration", 30) != liRenewalDaysBeforeExpiration )
                                        {
                                            moProfile["-RenewalDaysBeforeExpiration"] = liRenewalDaysBeforeExpiration;
                                            moProfile.Save();
                                        }

                                    ldtBeforeExpirationDate = aoOldCertificate.NotAfter.AddDays(-liRenewalDaysBeforeExpiration);
                                }

                    if ( DateTime.Now < ldtBeforeExpirationDate )
                    {
                        Env.LogIt(String.Format("Nothing to do until {0}. The current \"{1}\" certificate doesn't expire until {2}."
                            , ldtBeforeExpirationDate.ToString(), aoOldCertificate.Subject, aoOldCertificate.NotAfter.ToString()));

                        lbCertificateNotExpiring = true;
                    }
                }
            }

            return lbCertificateNotExpiring;
        }

        private static tvProfile oHandleUpdates(ref string[] args)
        {
            tvProfile   loProfile = null;
            tvProfile   loCmdLine = null;
            tvProfile   loMinProfile = null;
            string      lsRunKey = "-UpdateRunExePathFile";
            string      lsDelKey = "-UpdateDeletePathFile";
            string      lsUpdateRunExePathFile = null;
            string      lsUpdateDeletePathFile = null;
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
                    string lsFetchName = null;

                    //if ( null == DoGetCert.sProcessExePathFile(gsHostProcess) )
                    //{
                    //    // Fetch host (if it's not already running).
                    //    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    //            , String.Format("{0}{1}", Env.sFetchPrefix, gsHostProcess), loProfile.sRelativeToProfilePathFile(gsHostProcess));
                    //}

                    // Discard the default profile. Replace it with the WCF version fetched below.
                    File.Delete(loProfile.sLoadedPathFile);

                    // Fetch WCF config.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", Env.sFetchPrefix, "GetCert2.exe.config"), loProfile.sLoadedPathFile);

                    Env.ResetConfigMechanism(loProfile);

                    // Fetch security context task definition (for ADFS servers).
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="GetCert2Task.xml")
                            , loProfile.sRelativeToProfilePathFile(lsFetchName));
                }

                //if ( !loProfile.bExit && null != DoGetCert.sProcessExePathFile(gsHostProcess) )
                //{
                //    // Remove host (if it's already running).
                //    File.Delete(Path.Combine(Path.GetDirectoryName(loProfile.sLoadedPathFile), gsHostProcess));
                //}

                if ( loProfile.bExit || loProfile.bValue("-UseStandAloneMode", true) )
                    return loProfile;

                if ( !loProfile.bValue("-NoPrompts", false) )
                {
                    goStartupWaitMsg = new tvMessageBox();
                    goStartupWaitMsg.ShowWait(
                            null, Path.GetFileNameWithoutExtension(loProfile.sExePathFile) + " loading, please wait ...", 250);
                }

                loMinProfile = Env.oMinProfile(loProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);

                if ( String.IsNullOrEmpty(lsUpdateRunExePathFile) && String.IsNullOrEmpty(lsUpdateDeletePathFile) )
                {
                    // Check-in with the GetCert service.
                    loProfile.bExit = !DoGetCert.bDoCheckIn(lsHash, lbtArrayMinProfile);

                    // Short-circuit updates until basic setup is complete.
                    if ( "" == loProfile.sValue("-ContactEmailAddress", "") )
                        return loProfile;

                    // Does the "hosts" file need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleHostsEntryUpdate(loProfile, lsHash, lbtArrayMinProfile);

                    // A not-yet-updated EXE is running. Does it need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleExeUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile, lsRunKey, loCmdLine);
                }
                else
                {
                    // An updated EXE is running. It must be copied, relaunched then deleted.

                    if ( String.IsNullOrEmpty(lsUpdateDeletePathFile) )
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

                        Env.LogIt(String.Format(
                                "Software update successfully completed. Version {0} is now running.", lsCurrentVersion));
                    }
                }

                if ( !loProfile.bExit )
                {
                    loProfile = DoGetCert.oHandleIniUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile, args);

                    DoGetCert.HandleCfgUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile);
                    DoGetCert.HandleHostUpdates(loProfile, lsHash, lbtArrayMinProfile);

                    DoGetCert.HandleExeUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile, null, null);
                    DoGetCert.oHandleIniUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile, null);
                    DoGetCert.HandleCfgUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile);
                }
            }
            catch (Exception ex)
            {
                if ( null != loProfile )
                    loProfile.bExit = true;

                Env.LogIt(Env.sExceptionMessage(ex));

                if ( !loProfile.bValue("-UseStandAloneMode", true) && !loProfile.bValue("-CertificateSetupDone", false) )
                {
                    Env.LogIt(@"

Checklist for {EXE} setup:

    On this local server:

        .) Once ""-ContactEmailAddress"" and ""-CertificateDomainName"" have been
           provided (after the initial use of the ""{CertName}"" certificate),
           the SCS will expect to use a domain specific certificate for further
           communications. You can blank out ""-ContactEmailAddress"" in the profile
           file to force the use of the ""{CertName}"" certificate again
           (instead of an existing domain certificate).

        .) ""Host process can't be located on disk. Can't continue."" means the host
           software (ie. the task scheduler) must also be installed to continue.
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
                            .Replace("{CertName}", Env.sNewClientSetupCertName)
                            )
                            ;
                }

                try
                {
                    string lsDiscard = null;

                    DoGetCert.ReportErrors(lsHash, lbtArrayMinProfile, out lsDiscard);
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

        private static string sHostExePathFile(out Process aoProcessFound)
        {
            return DoGetCert.sHostExePathFile(out aoProcessFound, false);
        }
        private static string sHostExePathFile(out Process aoProcessFound, bool abQuietMode)
        {
            string  lsHostExeFilename = gsHostProcess;
            string  lsHostExePathFile = DoGetCert.sProcessExePathFile(lsHostExeFilename, out aoProcessFound);

            if ( String.IsNullOrEmpty(lsHostExePathFile) )
            {
                if ( !abQuietMode )
                    Env.LogIt("Host image can't be located on disk based on the currently running process. Trying typical locations ...");

                lsHostExePathFile = @"C:\ProgramData\GoPcBackup\GoPcBackup.exe";
                if ( !File.Exists(lsHostExePathFile) )
                    lsHostExePathFile = @"C:\Program Files\GoPcBackup\GoPcBackup.exe";

                if ( !File.Exists(lsHostExePathFile) )
                {
                    Env.LogIt("Host process can't be located on disk. Can't continue.");
                    throw new Exception("Exiting ...");
                }

                if ( !abQuietMode )
                    Env.LogIt("Host process image found on disk. It will be restarted.");
            }

            return lsHostExePathFile;
        }

        // Return an array of every site using the current setup certificate.
        private static Site[] oSetupCertBoundSiteArray(ServerManager aoServerManager)
        {
            if ( null == Env.oSetupCertificate )
                return null;

            List<Site>  loList = new List<Site>();
            byte[]      lbtSetupCertHashArray = Env.oSetupCertificate.GetCertHash();

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
            string  lsHostExePathFile = DoGetCert.sHostExePathFile(out loProcessFound, true);

            // Stop the EXE (can't update it or the INI while it's running).

            if ( !Env.bKillProcess(loProcessFound) )
            {
                Env.LogIt("Host process can't be stopped. Update can't be applied.");
                throw new Exception("Exiting ...");
            }
            else
            {
                Env.LogIt("Host process stopped successfully.");
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
                                loProfile = new tvProfile(aoProfile.sRelativeToProfilePathFile(lsExeName + ".exe.config"), true);
                                if ( loProfile.ContainsKey(lsKeyName) )
                                    lsVersion = loProfile.sValue(lsKeyName, "0");
                                break;
                        }
            string      lsIniUpdate = null;
                        // Look for profile update.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            switch (aeExeName)
                            {
                                case UpdatedEXEs.Client:
                                    lsIniUpdate = loGetCertServiceClient.sIniUpdate(asHash, abtArrayMinProfile, lsVersion);
                                    break;
                                case UpdatedEXEs.GcFailSafe:
                                    lsIniUpdate = loGetCertServiceClient.sFailSafeIniUpdate(asHash, abtArrayMinProfile, lsVersion);
                                    break;
                            }
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

            if ( String.IsNullOrEmpty(lsIniUpdate) )
            {
                Env.LogIt(String.Format("{0} profile version {1} is in use. No update.", lsExeName, lsVersion));
            }
            else
            {
                Env.LogIt(String.Format("{0} profile version {1} is in use. Update found ...", lsExeName, lsVersion));

                tvProfile loIniUpdateProfile = new tvProfile(lsIniUpdate);

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
                    aoProfile = new tvProfile(args, true);
                }

                Env.LogIt(String.Format(
                        "{0} profile update successfully completed. Version {1} is now in use.", lsExeName, loProfile.sValue("-IniVersion", "1")));
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
                                loProfile = new tvProfile(aoProfile.sRelativeToProfilePathFile(lsExeName + ".exe.config"), true);
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
                File.WriteAllText(loProfile.sLoadedPathFile, lsCfgUpdate);

                // Overwrite WCF config version number in the current profile.
                tvProfile   loWcfCfg = new tvProfile(loProfile.sLoadedPathFile, true);
                            loProfile["-CfgVersion"] = loWcfCfg.sValue("-CfgVersion", "1");
                            loProfile.Save();

                Env.LogIt(String.Format("{0} WCF conf update successfully completed. Version {1} is now in use.", lsExeName, loProfile.sValue("-CfgVersion", "1")));
            }
        }

        private static void HandleExeUpdate(UpdatedEXEs aeExeName, tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile, string asRunKey, tvProfile aoCmdLine)
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
                            lsUpdatePath = Path.Combine(Path.GetDirectoryName(lsExePathFile), aoProfile.sValue("-UpdateFolder", "Update"));
                            break;
                        case UpdatedEXEs.GcFailSafe:
                            lsExePathFile = aoProfile.sRelativeToProfilePathFile(lsExeName + ".exe");
                            lsIniPathFile = lsExePathFile + ".config";
                            lsUpdatePath = Path.GetDirectoryName(lsExePathFile);
                            break;
                    }
            string  lsCurrentVersion = !File.Exists(lsExePathFile) ? "0" : FileVersionInfo.GetVersionInfo(lsExePathFile).FileVersion;
            byte[]  lbtArrayExeUpdate = null;
                    // Look for EXE update.
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

            if ( null == lbtArrayExeUpdate )
            {
                Env.LogIt(String.Format("{0} version {1} is installed. No update.", lsExeName, lsCurrentVersion));
            }
            else
            {
                Env.LogIt("");
                Env.LogIt(String.Format("{0} version {1} is installed. Update found ...", lsExeName, lsCurrentVersion));

                string  lsNewExePathFile = Path.Combine(lsUpdatePath, Path.GetFileName(lsExePathFile));
                string  lsNewIniPathFile = Path.Combine(lsUpdatePath, Path.GetFileName(lsIniPathFile));

                Env.LogIt(String.Format("Writing update to \"{0}\".", lsNewExePathFile));
                Directory.CreateDirectory(lsUpdatePath);
                File.WriteAllBytes(lsNewExePathFile, lbtArrayExeUpdate);

                Env.LogIt(String.Format("The new file version of \"{0}\" is {1}."
                        , lsNewExePathFile, FileVersionInfo.GetVersionInfo(lsNewExePathFile).FileVersion));

                // Only the client needs this special handling (since it's running).
                if ( UpdatedEXEs.Client == aeExeName )
                {
                    Env.LogIt(String.Format("Writing update profile to \"{0}\".", lsNewIniPathFile));
                    File.Copy(lsIniPathFile, lsNewIniPathFile, true);

                    Env.LogIt(String.Format("Starting update \"{0}\" ...", lsNewExePathFile));
                    Process loProcess = new Process();
                            loProcess.StartInfo.FileName = lsNewExePathFile;
                            loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2}"
                                    , asRunKey, lsExePathFile, aoCmdLine.sCommandLine());
                            loProcess.Start();

                            System.Windows.Forms.Application.DoEvents();

                    aoProfile.bExit = true;
                }
            }
        }

        private static void HandleHostUpdates(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            bool        lbRestartHost = false;
            Process     loHostProcess = null;
            string      lsHostExePathFile= DoGetCert.sHostExePathFile(out loHostProcess);
            string      lsHostExeVersion = FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion;
            byte[]      lbtArrayGpcExeUpdate = null;
                        // Look for an updated GoPcBackup.exe.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lbtArrayGpcExeUpdate = loGetCertServiceClient.btArrayGoPcBackupExeUpdate(asHash, abtArrayMinProfile, lsHostExeVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( null != lbtArrayGpcExeUpdate )
                            Env.LogIt(String.Format("Host version {0} is in use. Update found ...", lsHostExeVersion));

            string      lsHostIniVersion = aoProfile.sValue("-HostIniVersion", "1");
            string      lsHostIniUpdate  = null;
                        // Look for an updated GoPcBackup.exe.txt.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lsHostIniUpdate = loGetCertServiceClient.sGpcIniUpdate(asHash, abtArrayMinProfile, lsHostIniVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( !String.IsNullOrEmpty(lsHostIniUpdate) )
                            Env.LogIt(String.Format("Host profile version {0} is in use. Update found ...", lsHostIniVersion));

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
                                            // Look for tasks in the host profile that are NOT in the update.
                                            foreach (DictionaryEntry loEntry in loHostProfile.oProfile("-AddTasks"))
                                            {
                                                tvProfile   loTask = new tvProfile(loEntry.Value.ToString());
                                                string      lsHostTaskEXE = loTask.sValue("-CommandEXE", "").ToLower();
                                                bool        lbAddTask = true;

                                                foreach (DictionaryEntry loEntry2 in loUpdateProfile.oProfile("-AddTasks"))
                                                {
                                                    tvProfile   loTaskUpdate = new tvProfile(loEntry2.Value.ToString());
                                                                if ( lsHostTaskEXE == loTaskUpdate.sValue("-CommandEXE", "").ToLower() )
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

        private static void HandleHostsEntryUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            string lsHstVersion = aoProfile.sValue("-HostsEntryVersion", "1");
            string lsHstUpdate = null;

            // Look for hosts entry update.
            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                lsHstUpdate = loGetCertServiceClient.sHostsEntryUpdate(asHash, abtArrayMinProfile, lsHstVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( String.IsNullOrEmpty(lsHstUpdate) )
            {
                Env.LogIt(String.Format("Hosts entry version {0} is in use. No update.", lsHstVersion));
            }
            else
            {
                Env.LogIt(String.Format("Hosts entry version {0} is in use. Update found ...", lsHstVersion));

                tvProfile       loHstUpdateProfile = new tvProfile(lsHstUpdate);
                string          lsIpAddress = loHstUpdateProfile.sValue("-IpAddress", "0.0.0.0");
                string          lsHostsPathFile = @"C:\windows\system32\drivers\etc\hosts";
                StringBuilder   lsbHostsStream = new StringBuilder(File.ReadAllText(lsHostsPathFile));
                string          lsDnsName = "secure_certificate_service_goes_here";
                Regex           loRegExRemoveComments = new Regex(String.Format(@"([#][0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)(.*)({0}.*\r\n)", lsDnsName), RegexOptions.IgnoreCase);
                                foreach (Match loMatch in loRegExRemoveComments.Matches(lsbHostsStream.ToString()))
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
                        Env.LogIt("Something's wrong with the local 'hosts' file format. Can't update IP address.");
                        Env.LogIt(Env.sExceptionMessage(ex));
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

                Env.LogIt(String.Format(
                        "Hosts entry update successfully completed. Version {0} is now in use.", aoProfile.sValue("-HostsEntryVersion", "1")));

                // Restart host (should restart this app as well - after a brief delay).
                Process loHostProcess = new Process();
                        loHostProcess.StartInfo.FileName = Path.Combine(Path.GetDirectoryName(DoGetCert.sStopHostExePathFile()), "Startup.cmd");
                        loHostProcess.StartInfo.Arguments = "-StartupTasksDelaySecs=10";
                        loHostProcess.StartInfo.UseShellExecute = false;
                        loHostProcess.StartInfo.CreateNoWindow = true;
                        loHostProcess.Start();

                Env.LogIt("Host process restarted successfully.");

                aoProfile.bExit = true;
            }
        }

        private static void VerifyDependenciesInit()
        {
            string  lsFetchName = null;

            // Fetch IIS Manager DLL.
            tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="Microsoft.Web.Administration.dll"), tvProfile.oGlobal().sRelativeToProfilePathFile(lsFetchName));
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
            byte[]              lbtArrayMinProfile = null;
            tvProfile           loMinProfile = null;
            X509Certificate2    loNewCertificate = null;
            X509Store           loStore = null;
            string              lsCertName = null;
            string              lsCertPathFile = null;
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

                lsCertPathFile = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-InstanceGuid", Guid.NewGuid().ToString()));
                loMinProfile = Env.oMinProfile(moProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);

                bool                lbSingleSessionEnabled = moProfile.bValue("-SingleSessionEnabled", false);
                string              lsDefaultPhysicalPath = moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot");
                bool                lbCertificatePassword = false;  // Set this "false" to use password only for load balancer file.
                string              lsCertificatePassword = HashClass.sHashPw(loMinProfile);
                bool                lbCreateSanSites = moProfile.bValue("-CreateSanSites", true);
                ServerManager       loServerManager = null;
                X509Certificate2    loOldCertificate = Env.oCurrentCertificate(lsCertName, out loServerManager);
                                    if ( this.bCertificateNotExpiring(lsHash, lbtArrayMinProfile, loOldCertificate) )
                                    {
                                        // The certificate is not ready to expire soon. There is nothing to do.
                                        
                                        return DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
                                    }
                byte[]              lbtArrayNewCertificate = null;
                string              lsCertOverridePfxComputer = DoGetCert.oDomainProfile.sValue("-CertOverridePfxComputer", "");
                string              lsCertOverridePfxPathFile = DoGetCert.oDomainProfile.sValue("-CertOverridePfxPathFile", "");
                bool                lbCertOverrideApplies = "" != lsCertOverridePfxPathFile && "" != lsCertOverridePfxComputer;
                bool                lbLoadBalancerCertPending = DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxPending", false);
                bool                lbRepositoryCertsOnly = DoGetCert.oDomainProfile.bValue("-RepositoryCertsOnly", false);
                DateTime            ldtMaintenanceWindowBeginTime = DoGetCert.oDomainProfile.dtValue("-MaintenanceWindowBeginTime", DateTime.MinValue);
                DateTime            ldtMaintenanceWindowEndTime   = DoGetCert.oDomainProfile.dtValue("-MaintenanceWindowEndTime",   DateTime.MaxValue);
                                    if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                    {
                                        if ( lbLoadBalancerCertPending )
                                        {
                                            int     liErrorCount = 1 + moProfile.iValue("-LoadBalancerPendingErrorCount", 0);
                                                    moProfile["-LoadBalancerPendingErrorCount"] = liErrorCount;
                                                    moProfile.Save();

                                            Env.LogIt("The load balancer administrator or process has not yet released the new certificate.");

                                            if ( liErrorCount > moProfile.iValue("-LoadBalancerPendingMaxErrors", 1) )
                                            {
                                                lbGetCertificate = false;

                                                Env.LogIt("This is an error.");
                                            }
                                            else
                                            {
                                                Env.LogIt("Will try again next cycle.");
                                                Env.LogIt("");

                                                return DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
                                            }
                                        }
                                        else
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
                                                if ( moProfile.bValue("-Auto", false) && (DateTime.Now < ldtMaintenanceWindowBeginTime || DateTime.Now >= ldtMaintenanceWindowEndTime) )
                                                {
                                                    lbGetCertificate = false;

                                                    Env.LogIt(String.Format("Can't proceed. \"{0}\" maintenance window is \"{1}\" to \"{2}\"."
                                                            , lsCertName, ldtMaintenanceWindowBeginTime, ldtMaintenanceWindowEndTime));
                                                }
                                                else
                                                {
                                                    Env.LogIt("New certificate downloaded from the repository.");

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

                    if ( liErrorCount > moProfile.iValue("-RepositoryCertsOnlyMaxErrors", 6) )
                    {
                        lbGetCertificate = false;

                        Env.LogIt("This is an error since -RepositoryCertsOnly=True.");
                    }
                    else
                    {
                        Env.LogIt("-RepositoryCertsOnly=True. Will try again next cycle.");
                        Env.LogIt("");

                        return DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && (!lbRepositoryCertsOnly || lbCertOverrideApplies) )
                {
                    if ( lbCertOverrideApplies )
                    {
                        Env.LogIt("");

                        if ( Env.sComputerName == lsCertOverridePfxComputer )
                        {
                            Env.LogIt(String.Format("A new certificate for \"{0}\" will come from the certificate override.", lsCertName));

                            lsCertPathFile = lsCertOverridePfxPathFile;
                        }
                        else
                        {
                            int liErrorCount = 1 + moProfile.iValue("-CertOverrideErrorCount", 0);
                                moProfile["-CertOverrideErrorCount"] = liErrorCount;
                                moProfile.Save();

                            Env.LogIt(String.Format("The certificate override for \"{0}\" will be found on another server (ie. \"{1}\").", lsCertName, lsCertOverridePfxComputer));

                            if ( liErrorCount > moProfile.iValue("-CertOverrideMaxErrors", 3) )
                            {
                                lbGetCertificate = false;

                                Env.LogIt("No more retries. This is an error.");
                                Env.LogIt("");
                            }
                            else
                            {
                                Env.LogIt("Will try again next cycle.");
                                Env.LogIt("");

                                return DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
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
                                    Env.LogIt(String.Format("Waiting at most {0} seconds to set a certificate renewal lock for this domain ...", liMaxCertRenewalLockDelaySecs));
                                    System.Windows.Forms.Application.DoEvents();
                                    Thread.Sleep(1000 * new Random().Next(liMaxCertRenewalLockDelaySecs));
                                }
                        bool    lbLockCertificateRenewal = false;
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
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
                                // certificate overrides are reserved to another computer on the domain).
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

                                    return DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
                                }

                        Env.LogIt("");
                        Env.LogIt(String.Format("Retrieving new certificate for \"{0}\" from the certificate provider network ...", lsCertName));

                        string  lsAcmePsFile = "ACME-PS.zip";
                        string  lsAcmePsPath = Path.GetFileNameWithoutExtension(lsAcmePsFile);
                        string  lsAcmePsPathFull = moProfile.sRelativeToProfilePathFile(lsAcmePsPath);
                        string  lsAcmePsPathFile = moProfile.sRelativeToProfilePathFile(lsAcmePsFile);
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

                            lsAcmePsPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsPath", lsAcmePsPath));

                            if ( String.IsNullOrEmpty(Path.GetDirectoryName(lsAcmePsPath)) )
                                lsAcmePsPath = Path.Combine(".", lsAcmePsPath);
                        }

                        string      lsAcmeWorkPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsWorkPath", "AcmeState"));
                        string      lsSanPsStringArray = "(";
                                    foreach (string lsSanItem in lsSanArray)
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

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
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
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
New-ACMENonce      ""{AcmeWorkPath}""
New-ACMEAccountKey ""{AcmeWorkPath}""
New-ACMEAccount    ""{AcmeWorkPath}"" -EmailAddresses ""{-ContactEmailAddress}"" -AcceptTOS -PassThru
New-ACMEOrder      ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                        }

                        ArrayList loAcmeCleanedList = new ArrayList();

                        for (int liSanArrayIndex=0; liSanArrayIndex < lsSanArray.Length; liSanArrayIndex++)
                        {
                            string lsSanItem = lsSanArray[liSanArrayIndex];

                            if ( lbGetCertificate && !this.bMainLoopStopped )
                            {
                                this.LogStage(String.Format("3 - Define DNS name to be challenged (\"{0}\"), setup domain challenge in IIS and submit it to certificate provider", lsSanItem));

                                Site    loSanSite = null;
                                Site    loPrimarySiteForDefaults = null;
                                        Env.oCurrentCertificate(lsSanItem, lsSanArray, out loServerManager, out loSanSite, out loPrimarySiteForDefaults);

                                // If -CreateSanSites is false, only create a default
                                // website (as needed) and use it for all SAN values.
                                if ( null == loSanSite && !lbCreateSanSites && 0 != loServerManager.Sites.Count )
                                {
                                    Env.LogIt(String.Format("No website found for \"{0}\". -CreateSanSites is \"False\". So no website created.\r\n", lsSanItem));

                                    loSanSite = loServerManager.Sites[0];
                                }

                                if ( null == loSanSite )
                                {
                                    string lsPhysicalPath = null;

                                    if ( 0 == loServerManager.Sites.Count )
                                    {
                                        Env.LogIt("No default website could be found.");

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

                                    Env.LogIt(String.Format("No website found for \"{0}\". New website created.", lsSanItem));
                                    Env.LogIt("");
                                }

                                string  lsAcmeBasePath = moProfile.sValue("-AcmeBasePath", Path.Combine(lsWellKnownBasePath, "acme-challenge"));
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
{AcmeSystemWide}
$challenge = Get-ACMEChallenge $global:state $global:authZ[$global:SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmeChallengePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge $global:state
                                                ")
                                                .Replace("{AcmePsPath}", lsAcmePsPath)
                                                .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                                .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                .Replace("{AcmeChallengePath}", lsAcmePath)
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

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge ""{AcmeWorkPath}""
                                                ")
                                                .Replace("{AcmePsPath}", lsAcmePsPath)
                                                .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
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

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
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
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(1000 * moProfile.iValue("-SubmissionWaitSecs", 10));

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

                            File.Delete(lsCertPathFile);

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
                                        .Replace("{CertificatePathFile}", lsCertPathFile)
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
                        Env.LogIt(Env.sExceptionMessage("The new certificate could not be found!"));

                    if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && !moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // Upload new certificate to the load balancer.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            if ( lbCertOverrideApplies )
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPathFile
                                                                                    , HashClass.sDecrypted(Env.oSetupCertificate, DoGetCert.oDomainProfile.sValue("-CertOverridePfxPassword", "")));
                            else
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPathFile, lsCertificatePassword);
                        }

                        // Upload new certificate to the repository.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lbGetCertificate = loGetCertServiceClient.bNewCertificateUploaded(lsHash, lbtArrayMinProfile, File.ReadAllBytes(lsCertPathFile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            Env.LogIt("");
                            if ( lbGetCertificate )
                                Env.LogIt("New certificate successfully uploaded to the repository.");
                            else
                                Env.LogIt(Env.sExceptionMessage("Failed uploading new certificate to repository."));
                        }
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped && null != lbtArrayNewCertificate || moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // A non-null "lbtArrayNewCertificate" means the certificate was just downloaded from the
                        // repository. Load it to be installed and bound locally (same if running stand-alone).
                        if ( !moProfile.bValue("-CertificatePrivateKeyExportable", false)
                                || ( !moProfile.bValue("-UseStandAloneMode", true) && !DoGetCert.oDomainProfile.bValue("-CertPrivateKeyExportAllowed", false)) )
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
                            Env.LogIt("");
                            Env.LogIt("The new certificate private key is exportable (\"-CertificatePrivateKeyExportable=True\"). This is not recommended.");

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

                    Env.LogIt("");
                    Env.LogIt(String.Format("Certificate thumbprint: {0} (\"{1}\")", loNewCertificate.Thumbprint, Env.sCertName(loNewCertificate)));

                    Env.LogIt("");
                    Env.LogIt("Install and bind certificate ...");

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
                        if ( DoGetCert.oDomainProfile.bValue("-NoIIS", false) )
                        {
                            Env.LogIt("");
                            Env.LogIt("Using non-IIS certificate binding ...");

                            string lsDiscard = null;
                            lbGetCertificate = Env.bRunPowerScript(out lsDiscard, gbMainLoopStopped, null, moProfile.sValue("-NoIISBindingScript", ""), false, true);
                        }
                        else
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            Env.LogIt("");
                            Env.LogIt("Bind certificate in IIS ...");

                            // Apply new site-to-cert bindings.
                            foreach (string lsSanItem in DoGetCert.sSanArrayAmended(loServerManager, lsSanArray))
                            {
                                if ( !lbGetCertificate || this.bMainLoopStopped )
                                    break;

                                Env.LogIt(String.Format("Applying new certificate to \"{0}\" site ...", lsSanItem));

                                Site    loSite = null;
                                        Env.oCurrentCertificate(lsSanItem, lsSanArray, out loServerManager, out loSite);

                                if ( null == loSite && lsSanItem == lsCertName )
                                {
                                    // No site found. Use the default site, if it exists (for the primary domain only).
                                    if ( 0 != loServerManager.Sites.Count )
                                        loSite = loServerManager.Sites[0];   // Default website.

                                    Env.LogIt(String.Format("No SAN website match could be found for \"{0}\". The default will be used (ie. \"{1}\").", lsSanItem, loSite.Name));
                                }

                                if ( null == loSite )
                                {
                                    // Still no site found (not even the default). This is an error.

                                    lbGetCertificate = false;

                                    Env.LogIt(String.Format("No website could be found to bind the new certificate for \"{0}\"\r\n(no default website either; FYI, -CreateSanSites is \"{1}\"). Can't continue.", lsSanItem, lbCreateSanSites));
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

                                            Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with no SNI) using \"{0}\".", lsNewBindingInformation));
                                        }
                                        else
                                        {
                                            // Set the SNI (Server Name Indication) flag.
                                            lsNewBindingInformation = String.Format(lsNewBindingInformation, lsSanItem);
                                            loNewBinding = loSite.Bindings.Add(lsNewBindingInformation, null == loOldCertificate ? null : loOldCertificate.GetCertHash(), loStore.Name);
                                            loNewBinding.SetAttributeValue("SslFlags", 1 /* SslFlags.Sni */);

                                            Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with SNI) using \"{0}\".", lsNewBindingInformation));
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        loSite.Bindings.Remove(loNewBinding);

                                        // Reverting to the default binding usage (must be an older OS).
                                        lsNewBindingInformation = "*:443:";
                                        loSite.Bindings.Add(lsNewBindingInformation, null == loOldCertificate ? null : loOldCertificate.GetCertHash(), loStore.Name);

                                        Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (on older OS) using \"{0}\".", lsNewBindingInformation));
                                    }

                                    loServerManager.CommitChanges();

                                    Env.LogIt(String.Format("Default SSL binding added to site: \"{0}\".", loSite.Name));
                                }

                                if ( lbGetCertificate && !this.bMainLoopStopped )
                                {
                                    bool lbBindingFound = false;

                                    foreach (Binding loBinding in loSite.Bindings)
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
                                                        lbDoBinding = (null != Env.oSetupCertificate && Env.oSetupCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash));
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
                                                Env.LogIt(String.Format("New certificate (\"{0}\") bound to port {1} in IIS for \"{2}\"."
                                                                , loNewCertificate.Thumbprint, lsBindingPort, lsSanItem));
                                            else
                                                Env.LogIt(String.Format("New certificate (\"{0}\") bound to address {1} and port {2} in IIS for \"{3}\"."
                                                                , loNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, lsSanItem));

                                            lbBindingFound = true;

                                            DoGetCert.WriteFailSafeFile(lsSanItem, loSite, loNewCertificate, lsWellKnownBasePath);
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

                                            Env.LogIt(String.Format("No SSL binding could be found for \"{0}\" (no default SSL binding either).\r\nCan't continue.", lsSanItem));
                                        }
                                        else
                                        {
                                            if ( null != loDefaultSslBinding.CertificateHash && loNewCertificate.GetCertHash().SequenceEqual(loDefaultSslBinding.CertificateHash) )
                                            {
                                                Env.LogIt(String.Format("Binding already applied to  \"{0}\"{1}", lsSanItem
                                                            , loSite.Name == lsSanItem ? "." : String.Format(" (via \"{0}\").", loSite.Name)));
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
                                                    Env.LogIt(String.Format("New certificate (\"{0}\") bound to port {1} in IIS for \"{2}\"."
                                                                    , loNewCertificate.Thumbprint, lsBindingPort, lsSanItem));
                                                else
                                                    Env.LogIt(String.Format("New certificate (\"{0}\") bound to address {1} and port {2} in IIS for \"{3}\"."
                                                                    , loNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, lsSanItem));

                                                DoGetCert.WriteFailSafeFile(lsSanItem, loSite, loNewCertificate, lsWellKnownBasePath);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if ( lbGetCertificate )
                        {
                            Env.LogSuccess();
                            Env.LogIt("");
                            Env.LogIt("A new certificate was successfully installed and bound.");
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
                                if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) )
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
                                    {
                                        // Remove the old cert.
                                        loStore.Remove(loOldCertificate);

                                        // Get old cert's private key file.
                                        lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, loOldCertificate);
                                        if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) )
                                        {
                                            // Remove old cert's private key file.
                                            File.Delete(lsMachineKeyPathFile);
                                        }

                                        Env.LogIt(String.Format("Old certificate (\"{0}\") removed from the local store.", loOldCertificate.Thumbprint));
                                    }
                                }
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && !moProfile.bValue("-UseStandAloneMode", true) )
                        {
                            // At this point we need to load the new certificate into the service factory object.
                            Env.oSetupCertificate = null;
                            Env.oGetCertServiceFactory = null;

                            Env.LogIt("Requesting removal of the old certificate from the repository.");

                            // Last step, remove old certificate from the repository (it can't be removed until the new certificate is in place everywhere).
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
                if ( "" == DoGetCert.oDomainProfile.sValue("-CertOverridePfxComputer", "") + DoGetCert.oDomainProfile.sValue("-CertOverridePfxPathFile", "") )
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

                if ( !lbGetCertificate && null != loNewCertificate)
                {
                    // Remove the new cert.
                    if ( null != loStore )
                        loStore.Remove(loNewCertificate);

                    // Get new cert's private key file.
                    string lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, loNewCertificate);
                    if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) )
                    {
                        // Remove new cert's private key file.
                        File.Delete(lsMachineKeyPathFile);
                    }
                }

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
                moProfile.Remove("-RepositoryCertsOnlyErrorCount");
                moProfile.Remove("-CertificateRenewalDateOverride");
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

                Env.LogIt("At least one stage failed (or the process was stopped). Check log for errors.");
            }

            Env.LogIt("");

            if ( lbGetCertificate )
                lbGetCertificate = DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);
            else
                DoGetCert.bDoCheckOut(lsHash, lbtArrayMinProfile);

            return lbGetCertificate;
        }
    }

    public enum UpdatedEXEs
    {
         Client = 1
        ,GcFailSafe = 2
    }
}
