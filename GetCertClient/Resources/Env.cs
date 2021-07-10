using Microsoft.Web.Administration;
using System;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using tvToolbox;

namespace GetCert2
{
    public class Env
    {
        public static event EventHandler<string> AppendOutputTextLine;  

        /// <summary>
        /// Returns the current certificate collection.
        /// </summary>
        public static X509Certificate2Collection oCurrentCertificateCollection
        {
            get
            {
                X509Certificate2Collection  loCurrentCertificateCollection = null;
                                            // Open the local machine certificate store (ie "Local Computer / Personal / Certificates").
                X509Store                   loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                                            loStore.Open(OpenFlags.ReadOnly);
                                            loCurrentCertificateCollection = loStore.Certificates;
                                            loStore.Close();

                return loCurrentCertificateCollection;
            }
        }

        /// <summary>
        /// This is the oGetCertServiceFactory object.
        /// </summary>
        public static ChannelFactory<GetCertService.IGetCertServiceChannel> oGetCertServiceFactory
        {
            get
            {
                if ( null == goGetCertServiceFactory )
                {
                    File.Delete(tvProfile.oGlobal().sRelativeToProfilePathFile(Env.sWcfLogFile));

                    goGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");

                    Env.SetCertificate();
                }

                return goGetCertServiceFactory;
            }
            set
            {
                goGetCertServiceFactory = value;
            }
        }
        private static ChannelFactory<GetCertService.IGetCertServiceChannel> goGetCertServiceFactory;

        public static tvProfile oMinProfile(tvProfile aoProfile)
        {
            tvProfile   loProfile = new tvProfile(aoProfile.ToString());
                        loProfile.Remove("-Help");
                        loProfile.Remove("-PreviousProcessOutputText");
                        loProfile.Add("-CurrentLocalTime", DateTime.Now);
                        loProfile.Add("-COMPUTERNAME", Env.sComputerName);

            return loProfile;
        }

        public static X509Certificate2 oSetupCertificate
        {
            get
            {
                return goSetupCertificate;
            }
            set
            {
                goSetupCertificate = value;
            }
        }
        private static X509Certificate2 goSetupCertificate = null;

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
        /// Returns the current server computer name.
        /// </summary>
        public static string sComputerName
        {
            get
            {
                return Env.oEnvProfile.sValue("-COMPUTERNAME", "Computer name not found.");
            }
        }

        /// <summary>
        /// Returns the current IIS port 443 bound certificate name (if any).
        /// </summary>
        public static string sCurrentCertificateName
        {
            get
            {
                string              lsCurrentCertificateName = null;
                X509Certificate2    loCurrentCertificate = Env.oCurrentCertificate();
                                    if ( null != loCurrentCertificate )
                                        lsCurrentCertificateName = Env.sCertName(loCurrentCertificate);

                return lsCurrentCertificateName;
            }
        }

        /// <summary>
        /// Returns the current "LogPathFile" name.
        /// </summary>
        public static string sLogPathFile
        {
            get
            {
                return Env.sUniqueIntOutputPathFile(
                            Env.sLogPathFileBase
                        , tvProfile.oGlobal().sValue("-LogFileDateFormat", "-yyyy-MM-dd")
                        , true
                        );
            }
        }

        public static string sNewClientSetupCertName
        {
            get
            {
                return gsNewClientSetupCertName;
            }
        }
        private static string    gsNewClientSetupCertName  = "GetCertClientSetup";


        private static tvProfile oEnvProfile
        {
            get
            {
                if ( null == goEnvProfile )
                {
                    goEnvProfile = new tvProfile();

                    foreach (DictionaryEntry loEntry in Environment.GetEnvironmentVariables())
                    {
                        goEnvProfile.Add("-" + loEntry.Key.ToString(), loEntry.Value);
                    }
                }

                return goEnvProfile;
            }
        }
        private static tvProfile goEnvProfile = null;

        /// <summary>
        /// Returns the current "LogPathFile" base name.
        /// </summary>
        private static string sLogPathFileBase
        {
            get
            {
                string lsPath = "Logs";
                string lsLogFile = "Log.txt";
                string lsLogPathFileBase = lsLogFile;

                try
                {
                    lsLogPathFileBase = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-LogPathFile"
                            , Path.Combine(lsPath, Path.GetFileNameWithoutExtension(tvProfile.oGlobal().sLoadedPathFile) + lsLogFile)));

                    lsPath = Path.GetDirectoryName(lsLogPathFileBase);
                    if ( !Directory.Exists(lsPath) )
                        Directory.CreateDirectory(lsPath);
                }
                catch {}

                return lsLogPathFileBase;
            }
        }

        // This kludge is necessary since neither "Process.CloseMainWindow()" nor "Process.Kill()" work reliably.
        public static bool bKillProcess(Process aoProcess)
        {
            bool    lbKillProcess = true;
                    if ( null == aoProcess )
                        return lbKillProcess;

            int     liKillProcessOrderlyWaitSecs = 10;
            int     liKillProcessForcedWaitMS = 1000;
                    if ( null != tvProfile.oGlobal() )
                    {
                        liKillProcessOrderlyWaitSecs = tvProfile.oGlobal().iValue("-KillProcessOrderlyWaitSecs", liKillProcessOrderlyWaitSecs);
                        liKillProcessForcedWaitMS = tvProfile.oGlobal().iValue("-KillProcessForcedWaitMS", liKillProcessForcedWaitMS);
                    }
            int     liProcessId = 0;

            // First try ending the process in the usual way (ie. an orderly shutdown).

            Process loKillProcess = new Process();

            try
            {
                if ( null != aoProcess )
                    liProcessId = aoProcess.Id;

                loKillProcess = new Process();
                loKillProcess.StartInfo.FileName = "taskkill";
                loKillProcess.StartInfo.Arguments = "/pid " + liProcessId;
                loKillProcess.StartInfo.UseShellExecute = false;
                loKillProcess.StartInfo.CreateNoWindow = true;
                loKillProcess.Start();

                if ( null != aoProcess )
                    aoProcess.WaitForExit(1000 * liKillProcessOrderlyWaitSecs);
            }
            catch (Exception ex)
            {
                lbKillProcess = false;
                Env.LogIt(String.Format("Orderly bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, Env.sExceptionMessage(ex)));
            }

            if ( lbKillProcess && null != aoProcess && !aoProcess.HasExited )
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

                    if ( null != aoProcess )
                        aoProcess.WaitForExit(liKillProcessForcedWaitMS);

                    lbKillProcess = null != aoProcess || aoProcess.HasExited;
                }
                catch (Exception ex)
                {
                    lbKillProcess = false;
                    Env.LogIt(String.Format("Forced bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, Env.sExceptionMessage(ex)));
                }
            }

            return lbKillProcess;
        }

        public static bool bKillProcessParent(Process aoProcess)
        {
            bool lbKillProcess = true;

            foreach (Process loProcess in Process.GetProcessesByName(aoProcess.ProcessName))
            {
                if ( loProcess.Id != aoProcess.Id )
                    lbKillProcess = Env.bKillProcess(loProcess);
            }

            return lbKillProcess;
        }

        public static bool bRunPowerScript(out string asOutput, string asScript)
        {
            return Env.bRunPowerScript(out asOutput, false, null, asScript, false, false);
        }
        public static bool bRunPowerScript(out string asOutput, bool abMainLoopStopped, string asSingleSessionScriptPathFile, string asScript, bool abOpenOrCloseSingleSession, bool abSkipLog)
        {
            bool            lbSingleSessionEnabled = tvProfile.oGlobal().bValue("-SingleSessionEnabled", false);
                            if ( !lbSingleSessionEnabled && abOpenOrCloseSingleSession )
                            {
                                // -SingleSessionEnabled is false, so ignore the single session "open" and "close" scripts.
                                asOutput = null;
                                return true;
                            }
            bool            lbRunPowerScript = false;
            string          lsScriptPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PowerScriptPathFile", "PowerScript.ps1"));
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

                            File.WriteAllText(lsScriptPathFile, lsScript);
                            System.Windows.Forms.Application.DoEvents();
                            Thread.Sleep(tvProfile.oGlobal().iValue("-PowerScriptSleepMS", 200));

                            if ( File.ReadAllText(lsScriptPathFile) != lsScript )
                            {
                                Env.LogIt(Env.sExceptionMessage("What was written to disk does not match the given snippet. Can't continue."));
                                asOutput = null;
                                return false;
                            }
            string  lsProcessPathFile = tvProfile.oGlobal().sValue("-PowerShellExePathFile", @"C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe");
            string  lsProcessArgs = tvProfile.oGlobal().sValue("-PowerShellExeArgs", @"-NoProfile -ExecutionPolicy unrestricted -File ""{0}"" ""{1}""");
                    if ( null == asSingleSessionScriptPathFile )
                        lsProcessArgs = String.Format(lsProcessArgs, lsScriptPathFile, "");
                    else
                        lsProcessArgs = String.Format(lsProcessArgs, asSingleSessionScriptPathFile, lsScriptPathFile);
            string  lsLogScript = !lsScript.Contains("-CertificateKey") ? lsScript : lsScript.Substring(0, lsScript.IndexOf("-CertificateKey"));

            if ( !abSkipLog && (!lbSingleSessionEnabled || !abOpenOrCloseSingleSession) )
                Env.LogIt(lsLogScript);

            Env.bPowerScriptSkipLog = abSkipLog;

            Process loProcess = new Process();
                    loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                    loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);
                    loProcess.StartInfo.FileName = lsProcessPathFile;
                    loProcess.StartInfo.Arguments = lsProcessArgs;
                    loProcess.StartInfo.UseShellExecute = false;
                    loProcess.StartInfo.RedirectStandardError = true;
                    loProcess.StartInfo.RedirectStandardInput = true;
                    loProcess.StartInfo.RedirectStandardOutput = true;
                    loProcess.StartInfo.CreateNoWindow = true;

            System.Windows.Forms.Application.DoEvents();

            Env.bPowerScriptError = false;
            Env.sPowerScriptOutput = null;

            loProcess.Start();
            loProcess.BeginErrorReadLine();
            loProcess.BeginOutputReadLine();

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(tvProfile.oGlobal().dValue("-PowerScriptTimeoutSecs", 300));

            // Wait for the process to finish.
            while ( !abMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !abMainLoopStopped )
                    Thread.Sleep(tvProfile.oGlobal().iValue("-PowerScriptSleepMS", 200));
            }

            if ( !abMainLoopStopped )
            {
                if ( !loProcess.HasExited )
                    Env.LogIt(tvProfile.oGlobal().sValue("-PowerScriptTimeoutMsg", "*** sub-process timed-out ***\r\n\r\n"));

                int liExitCode = -1;
                    try { liExitCode = loProcess.ExitCode; } catch {}

                if ( Env.bPowerScriptError || liExitCode != 0 || !loProcess.HasExited )
                {
                    Env.LogIt(Env.sExceptionMessage("The sub-process experienced a critical failure."));
                }
                else
                {
                    lbRunPowerScript = true;

                    if ( !String.IsNullOrEmpty(lsScriptPathFile) )
                        File.Delete(lsScriptPathFile);

                    if ( !abSkipLog && (!lbSingleSessionEnabled || !abOpenOrCloseSingleSession) )
                        Env.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            Env.bKillProcess(loProcess);

            asOutput = Env.sPowerScriptOutput;
            Env.bPowerScriptSkipLog = false;

            return lbRunPowerScript;
        }
        public static void PowerScriptProcessOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length )
            {
                Env.sPowerScriptOutput += outLine.Data;

                if ( !Env.bPowerScriptSkipLog )
                    Env.LogIt(outLine.Data);
            }
        }
        public static void PowerScriptProcessErrorHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length )
            {
                Env.sPowerScriptOutput += outLine.Data;
                Env.bPowerScriptError = true;
                Env.LogIt(outLine.Data);
            }
        }

        /// <summary>
        /// Returns the current IIS port 443 bound certificate (if any), otherwise whatever's in the local store (by name).
        /// </summary>
        public static X509Certificate2 oCurrentCertificate()
        {
            return Env.oCurrentCertificate(null);
        }
        private static X509Certificate2 oCurrentCertificate(string asCertName)
        {
            string[]        lsSanArray = null;
            ServerManager   loServerManager = null;
            Site            loSite = null;
            Site            loPrimarySiteForDefaults = null;

            return Env.oCurrentCertificate(asCertName, lsSanArray, out loServerManager, out loSite, out loPrimarySiteForDefaults);
          }
        public static X509Certificate2 oCurrentCertificate(string asCertName, out ServerManager aoServerManager)
        {
            string[]        lsSanArray = null;
            Site            loSite = null;
            Site            loPrimarySiteForDefaults = null;

            return Env.oCurrentCertificate(asCertName, lsSanArray, out aoServerManager, out loSite, out loPrimarySiteForDefaults);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out Site aoSiteFound
                )
        {
            ServerManager   loServerManager = null;
            Site            loPrimarySiteForDefaults = null;

            return Env.oCurrentCertificate(asCertName, asSanArray, out loServerManager, out aoSiteFound, out loPrimarySiteForDefaults);
        }
        public static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                )
        {
            Site            loPrimarySiteForDefaults = null;

            return Env.oCurrentCertificate(asCertName, asSanArray, out aoServerManager, out aoSiteFound, out loPrimarySiteForDefaults);
        }
        public static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , out Site aoPrimarySiteForDefaults
                )
        {
            X509Certificate2    loCurrentCertificate = null;

            // ServerManager is the IIS ServerManager. It gives us the website object for binding the cert.
            aoServerManager = new ServerManager();
            aoSiteFound = null;
            aoPrimarySiteForDefaults = null;

            try
            {
                string  lsCertName = null == asCertName ? null : asCertName.ToLower();
                Binding loBindingFound = null;

                foreach (Site loSite in aoServerManager.Sites)
                {
                    // First, look for a binding by matching certificate.
                    foreach (Binding loBinding in loSite.Bindings)
                    {
                        // Select the first binding with a certificate (and a matching name - if lsCertName is not null).
                        foreach (X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                        {
                            if ( null != loBinding.CertificateHash && loCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash)
                                && (String.IsNullOrEmpty(lsCertName) || lsCertName == Env.sCertName(loCertificate).ToLower()) )
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
                            if ( null == loBinding.CertificateHash && !String.IsNullOrEmpty(lsCertName)
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

                if ( null == aoSiteFound && !String.IsNullOrEmpty(lsPrimaryCertName) && lsCertName != lsPrimaryCertName )
                {
                    // Use the primary site in the SAN array as defaults for any new site (to be created).
                    Env.oCurrentCertificate(lsPrimaryCertName, asSanArray, out aoPrimarySiteForDefaults);
                }

                // As a last resort (so much for finally), use the newest certificate found in the local store (by name).
                if ( null == loCurrentCertificate && !String.IsNullOrEmpty(lsCertName) )
                    foreach (X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                    {
                        if ( lsCertName == Env.sCertName(loCertificate).ToLower()
                                && (null == loCurrentCertificate || loCertificate.NotAfter > loCurrentCertificate.NotAfter) )
                        {
                            loCurrentCertificate = loCertificate;
                        }
                    }
            }
            catch (Exception ex)
            {
                Env.LogIt(ex.Message);
                Env.LogIt("IIS may not be fully configured. This will prevent certificate-to-site binding in IIS.");
            }

            return loCurrentCertificate;
        }

        public static string sExceptionMessage(Exception ex)
        {
            return String.Format("ServiceFault: {0}{1}{2}{3}", ex.Message, (null == ex.InnerException ? "": "; " + ex.InnerException.Message), Environment.NewLine, ex.StackTrace);
        }
        public static string sExceptionMessage(string asMessage)
        {
            return "ServiceFault: " + asMessage;
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

        public static bool      bPowerScriptError;
        public static bool      bPowerScriptSkipLog;
        public static string    sPowerScriptOutput;
        public static string    sFetchPrefix = "Resources.Fetch.";
        public static string    sWcfLogFile = "WcfLog.txt";         // Must be changed in the WCF config too.

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
                lsLogPathFileBase = Env.sLogPathFileBase;

                // Move down to the base folder if we're running an update.
                string  lsPathSep = Path.DirectorySeparatorChar.ToString();
                string  lsPath = Path.GetDirectoryName(tvProfile.oGlobal().sRelativeToProfilePathFile(lsLogPathFileBase)).Replace(
                                    String.Format("{0}{1}{0}",  lsPathSep, tvProfile.oGlobal().sValue("-UpdateFolder", "Update")), lsPathSep);
                        Directory.CreateDirectory(lsPath);

                loStreamWriter = new StreamWriter(tvProfile.oGlobal().sRelativeToProfilePathFile(
                        Env.sUniqueIntOutputPathFile(lsLogPathFileBase
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

            if ( null != Env.AppendOutputTextLine )
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke((Action)(() =>
            {
                Env.AppendOutputTextLine(null, asText);
                System.Windows.Forms.Application.DoEvents();
            }
            ));
        }

        public static void LogDone()
        {
            Env.LogIt("Done.");
        }

        public static void LogSuccess()
        {
            Env.LogIt("Success.");
        }

        // Reset WCF configuration, from "https://stackoverflow.com/questions/6150644/change-default-app-config-at-runtime".
        public static void ResetConfigMechanism(tvProfile aoProfile)
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

        /// <summary>
        /// Set client certificate.
        /// </summary>
        public static void SetCertificate()
        {
            tvProfile loProfile = tvProfile.oGlobal();

            if ( loProfile.bValue("-UseStandAloneMode", true) || loProfile.bExit )
                return;

            if ( null == goGetCertServiceFactory.Credentials.ClientCertificate.Certificate
                    || (loProfile.bValue("-CertificateSetupDone", false) && goGetCertServiceFactory.Credentials.ClientCertificate.Certificate.Subject.Contains("*")) )
            {
                if ( null == goSetupCertificate )
                {
                    try
                    {
                        if ( "" != loProfile.sValue("-ContactEmailAddress" ,"") )
                            goSetupCertificate = Env.oCurrentCertificate(loProfile.sValue("-CertificateDomainName" ,""));

                        if ( null == goSetupCertificate && !loProfile.bValue("-CertificateSetupDone", false) && !loProfile.bValue("-NoPrompts", false) )
                        {
                            System.Windows.Forms.OpenFileDialog loOpenDialog = new System.Windows.Forms.OpenFileDialog();
                                                                loOpenDialog.FileName = gsNewClientSetupCertName + ".pfx";
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

                                        Env.LogIt(Env.sExceptionMessage(ex));
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        loProfile.bExit = true;

                        Env.LogIt(Env.sExceptionMessage(ex));

                        throw ex;
                    }
                }

                goGetCertServiceFactory.Credentials.ClientCertificate.Certificate = goSetupCertificate;
            }

            if ( null == goSetupCertificate )
                goSetupCertificate = goGetCertServiceFactory.Credentials.ClientCertificate.Certificate;
        }
    }
}
