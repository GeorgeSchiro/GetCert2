Overview
========


**GetCert2** is simple and **FREE** software for automating digital certificate retrieval and installation (screenshots below).

This utility gets a digital certificate from the **FREE** "Let's Encrypt" certificate provider network (see 'LetsEncrypt.org'). It installs the certificate in your server's local computer certificate store and binds it to port 443 in IIS.

If the current time is not within a given number of days prior to expiration of the current digital certificate (eg. 30 days, see -RenewalDaysBeforeExpiration below), this software does nothing. Otherwise, the retrieval process begins. If 'stand-alone' mode is disabled (see -UseStandAloneMode below), the certificate retrieval process is used in concert with the secure certificate service (SCS, see 'GoGetCert.com').

If the software is not running in 'stand-alone' mode, it also copies any new cert to a secure file anywhere on the local area network to be picked up by the load balancer administrator or process. It also replaces the SSO (single-sign-on) certificate in your central SSO configuration (eg. ADFS) and restarts the SSO service on all servers in any defined SSO server farm. It also replaces all integrated application SSO certificate references in any number of configuration files anywhere on the local network.

Also when -UseStandAloneMode is False, this software checks for any previously fetched certificate (by another client running the same software). If found, it is downloaded directly from the SCS (rather than attempting retrieval from the certificate provider network). If not found on the SCS, a certificate is requested from the certificate provider network, uploaded to the SCS (for use by other clients in the same server farm) and installed locally and bound to port 443 in IIS. Finally, calls to the certificate provider network can be overridden by providing this software access to a secure digital certificate file anywhere on the local network.

Give it a try.

The first time you run 'GetCert2.exe' it will prompt you to create a 'GetCert2' folder on your desktop. It will copy itself and continue running from there.

Everything the software needs is written to the 'GetCert2' folder. Nothing is written anywhere else (except IIS configuration and your certificates).

If you like the software, you can leave it on your desktop or you can run 'Setup Application Folder.exe' (also in the 'GetCert2' folder, be sure to run it as administrator). If you decide not to keep the software, simply delete the 'GetCert2' folder from your desktop.


Features
========


-   Simple setup - try it out fast
-   Only two input details needed to get started
-   Performs extensive handshaking with the "Let's Encrypt"
    certificate provider network, all completely automatically
-   Installs a new certificate and binds it to port 443 in IIS
-   Creates new websites in IIS (as needed) for each SAN domain
-   Comprehensive dated log files are produced with every run
-   Can be command-line driven from a server batch job scheduler
-   Software is highly configurable
-   Software is totally self-contained (EXE is its own setup)


**GetCert2** is essentially an automation front-end for 'ACME-PS'. 'ACME-PS' is an excellent tool. That said, you can replace it with any other PowerShell capable ACME protocol tool you might prefer instead. Such a change would be made in the profile file like everything else (see -AcmePsPath, -ScriptStage1, etc. below).

Note: since wildcard (ie. star) certificates are no longer considered 'security best-practice' (see [NSA: Avoid Dangers of Wildcard TLS Certificates](https://www.nsa.gov/Press-Room/News-Highlights/Article/Article/2804293/avoid-dangers-of-wildcard-tls-certificates-the-alpaca-technique/), [What vulnerabilities could be caused by a wildcard SSL cert?](https://security.stackexchange.com/questions/8210/what-vulnerabilities-could-be-caused-by-a-wildcard-ssl-cert)), **GetCert2** doesn't include support for them.

Note: anything from the profile (see 'Options and Features' below) can be passed thru any -ScriptStage snippet to PowerShell as a string token of the form: {-KeyName}. Here's an example:

    New-ACMEAccount $state -EmailAddresses "{-ContactEmailAddress}" -AcceptTOS


Screenshots
===========


![Copy EXE to Desktop?](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot0.PNG)
![Setup Wizard        ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot1.PNG)
![Setup Wizard Step 1 ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot2.PNG)
![Setup Wizard Done   ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot3.PNG)
![Run Process Prompt  ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot4.PNG)
![Process Running 1   ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot5.PNG)
![Process Running 2   ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot6.PNG)
![Process Finished    ](https://raw.githubusercontent.com/GeorgeSchiro/GetCert2/master/Project/Screenshots/Shot7.PNG)


Requirements
============


-   Windows Server 2016+
-   Internet Information Services (IIS)
-   .Net Framework 4.8+


Command-Line Usage
==================


    Open this utility's profile file to see additional options available. It is 
    usually located in the same folder as 'GetCert2.exe' and has the same name 
    with '.config' added (see 'GetCert2.exe.config').

    Profile file options can be overridden with command-line arguments. The keys 
    for any '-key=value' pairs passed on the command-line must match those that 
    appear in the profile (with the exception of the '-ini' key).

    For example, the following invokes the use of an alternative profile file:
    (be sure to copy an existing profile file if you do this):

        GetCert2.exe -ini=NewProfile.txt

    This tells the software to run in automatic mode:

        GetCert2.exe -Auto


    Author:  George Schiro

    Date:    10/24/2019


Options and Features
====================


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
    "Let's Encrypt" certificate network. By default, this key is set to the 'ACME-PS'
    folder which, with no absolute path given, will be expected to be found within
    the folder containing 'GetCert2.exe.config'. Set -AcmePsModuleUseGallery=True
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

-DnsChallengeSleepSecs=15

    This is the number of seconds to wait after updating the global DNS
    database with new domain specific challenge data and before submitting
    the ACME challenge to the certificate provider network.

-DoStagingTests=True

    Initial testing is done with the certificate provider staging network.
    Set this switch False to use the live production certificate network.

-FetchSource=False

    Set this switch True to fetch the source code for this utility from the EXE.
    Look in the containing folder for a ZIP file with the full project sources.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-KeysAfterReset= SEE PROFILE FOR DEFAULT VALUE

    This is the list of profile keys that are preserved whenever the SCS method
    'ResetConfig' is used (see 'GoGetCert Client Callable SCS Methods.pdf').
    If you've made customizations, be sure to add the changed keys to this list.

    Note: this parameter is ignored when -UseStandAloneMode is True.

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

    Set this switch True to see in detail how the GetCert2.exe app's identifying
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

-PowerShellExeArgs=-NoProfile -ExecutionPolicy unrestricted -File "{0}" "{1}"

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

-ResetAccountFilesEachRun=False

    This switch allows you to override the default behavior of only removing
    certificate provider network account files when changing between testing and
    production mode. These account files are automatically created if they don't
    already exist. It is recommended that you preserve these account files (see
    -AcmeAccountFile and -AcmeAccountKeyFile above) if you are doing very high
    volumes of ACME DNS challenges.

-ResetStagingLogs=False

    Having to wade through several log sessions during testing can be cumbersome.
    Setting this switch True will clear previous log sessions on the client after
    each staging test.

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
    in the GetCert2.exe UI. Click the 'SAN List' button to see the proper format here.

    Here's a command-line example:

    -SanList="-Domain='MyDomain.com' -Domain='www.MyDomain.com'"

    Note: the GetCert2.exe UI limits you to 100 SAN values (the certificate provider
          does the same). If you add more names than this limit, an error results and
          no certificate will be generated.

-SaveProfile=True

    Set this switch False to prevent saving to the profile file by this utility
    software itself. This is not recommended since process status information is 
    written to the profile after each run.

-SaveSansCmdLine=True

    Set this switch False to allow merged command-lines to be written to
    the profile file (ie. 'GetCert2.exe.config'). When True, everything
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

-ScriptStage3Dns= SEE PROFILE FOR DEFAULT VALUE

    This is where you customize adding TXT records to the global DNS database for
    the purpose of doing ACME DNS challenges. The details are specific to your DNS
    provider (which must allow for automated changes to your DNS records).

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

        GetCert2.exe -Auto -Setup

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
    client only). See -SsoThumbprintFiles below.

-SsoMaxRenewalSleepSecs=60

    This is the maximum seconds an SSO server will wait for another
    SSO server (in the same farm) to renew the SSO certificates.

-SsoProxySleepSecs=60

    An SSO proxy server will wait this many seconds before
    polling the SSO server again for a certificate change.

-SsoProxyTimeoutMins=30

    An SSO proxy server will wait this many minutes for the SSO server
    to change its certificates before throwing a timeout exception.

-SsoThumbprint= NO DEFAULT VALUE

    This is the thumbprint value of the current SSO certificate. It is received
    from the SCS. When it changes, corresponding changes must also be made to
    application related configuration files (see -SsoThumbprintFiles below).

-SsoThumbprintFiles='C:\inetpub\wwwroot\web.config'

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

-UseDnsChallenge=False

    The default ACME challenge type is the HTTP challenge. Set this switch True
    if you will be doing ACME DNS challenges instead. DNS challenges are needed
    for large enterprises with many internal domains (ie. domains that you'd
    rather not add to the global DNS database).

-UseNonIISBindingAlso=False

    Set this switch True to use the typical IIS binding procedures
    on this machine as well as the -NonIISBindingScript (see above).

-UseNonIISBindingOnly=False

    Set this switch True to disable the usual IIS binding procedures on
    this machine and use the -NonIISBindingScript instead (see above).

-UseNonIISBindingPfxFile=False

    Set this switch True to allow for the creation of a PFX file for use by
    the non-IIS binding script (see -NonIISBindingScript above). The server
    location and the password of this PFX file must be defined on the SCS
    (see 'GoGetCert Client Callable SCS Methods.pdf').

-UseStandAloneMode=True

    Set this switch False and the software will use the GoGetCert Secure Certificate
    Service (see 'GoGetCert.com') to manage certificates between several servers
    in a server farm, on SSO servers, SSO integrated application servers and load
    balancers.


Notes:

    There may be various other settings that can be adjusted also (user
    interface settings, etc). See the profile file ('GetCert2.exe.config')
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added 'on the fly'
    (in order of execution) to 'GetCert2.exe.config' as the software runs.
