using Microsoft.Web.Administration;
using System;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security.Cryptography;
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
        /// Returns the current "non-error bounce" status.
        /// </summary>
        public static bool bCanScheduleNonErrorBounce
        {
            get
            {
                return !File.Exists(Env.sBounceOnNonErrorLockPathFile);
            }
        }


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

                    try
                    {
                        goGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                    }
                    catch (Exception ex)
                    {
                        Env.LogIt(String.Format("Failure occurred attempting to open the service communication channel.{0}{1}"
                                                , Environment.NewLine, Env.sExceptionMessage(ex)));
                    }

                    if ( null != goGetCertServiceFactory )
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

        public static string sBounceOnErrorLockPathFile
        {
            get
            {
                if ( null == msBounceOnErrorLockPathFile )
                    msBounceOnErrorLockPathFile = tvProfile.oGlobal().sRelativeToExePathFile("BounceOnErrorScheduled.lck");

                return msBounceOnErrorLockPathFile;
            }
        }
        private static string msBounceOnErrorLockPathFile = null;

        public static string sBounceOnNonErrorLockPathFile
        {
            get
            {
                if ( null == msBounceOnNonErrorLockPathFile )
                    msBounceOnNonErrorLockPathFile = tvProfile.oGlobal().sRelativeToExePathFile("BounceOnNonErrorScheduled.lck");

                return msBounceOnNonErrorLockPathFile;
            }
        }
        private static string msBounceOnNonErrorLockPathFile = null;

        public static X509Certificate2 oChannelCertificate
        {
            get
            {
                return goChannelCertificate;
            }
            set
            {
                goChannelCertificate = value;
            }
        }
        private static X509Certificate2 goChannelCertificate = null;

        public static tvProfile oDomainProfile
        {
            get
            {
                return goDomainProfile;
            }
            set
            {
                goDomainProfile = value;

                if ( null != goDomainProfile )
                {
                    // Any star cert name must be cached since a non-empty value
                    // is needed at startup (before the domain profile is loaded).

                    string lsStarCertName = goDomainProfile.sValue(Env.sStarCertNameKey, "");

                    if ( "" != lsStarCertName )
                    {
                        if ( lsStarCertName != tvProfile.oGlobal().sValue(Env.sStarCertNameKey, "") )
	                    {
                            tvProfile.oGlobal()[Env.sStarCertNameKey] = lsStarCertName;
                            tvProfile.oGlobal().Save();
	                    }
                    }
                    else
	                {
                        if ( tvProfile.oGlobal().ContainsKey(Env.sStarCertNameKey) )
                        {
                            tvProfile.oGlobal().Remove(Env.sStarCertNameKey);
                            tvProfile.oGlobal().Save();
                        }
	                }
                }
            }
        }
        private static tvProfile goDomainProfile = null;

        public static tvProfile oMinProfile(tvProfile aoProfile)
        {
            if ( null == goMinProfile )
            {
                goMinProfile = new tvProfile(aoProfile.ToString());
                goMinProfile.Remove("-Help");
                goMinProfile.Remove("-PreviousProcessOutputText");
                goMinProfile.Add("-CurrentLocalTime", DateTime.Now);
                goMinProfile.Add("-COMPUTERNAME", Env.sComputerName);
            }

            return goMinProfile;
        }
        private static tvProfile goMinProfile = null;

        /// <summary>
        /// Returns the certificate name from a certificate.
        /// </summary>
        /// <param name="aoCertificate"></param>
        public static string sCertName(X509Certificate2 aoCertificate)
        {
            return null == aoCertificate ? null : aoCertificate.GetNameInfo(X509NameType.SimpleName, false);
        }
        public static string sCertName(string asCertNameOrSanItem)
        {
            tvProfile   loDomainProfile = null != Env.oDomainProfile ? Env.oDomainProfile : new tvProfile();
            string      lsCertName = loDomainProfile.sValue(Env.sStarCertNameKey, tvProfile.oGlobal().sValue(Env.sStarCertNameKey, ""));

            return null != asCertNameOrSanItem && "" != lsCertName ? lsCertName : asCertNameOrSanItem;
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
                          tvProfile.oGlobal().sRelativeToProfilePathFile(Env.sLogPathFileBase)
                        , tvProfile.oGlobal().sValue("-LogFileDateFormat", "-yyyy-MM-dd")
                        , true
                        );
            }
        }


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
                string  lsLogPathFileBase = tvProfile.oGlobal().sValue("-LogPathFile"
                                , Path.Combine("Logs", Path.GetFileNameWithoutExtension(tvProfile.oGlobal().sLoadedPathFile) + "Log.txt"));
                        lsLogPathFileBase = Path.Combine(Path.GetDirectoryName(lsLogPathFileBase), String.Format("{0}({1}){2}"
                                , Path.GetFileNameWithoutExtension(lsLogPathFileBase), tvProfile.oGlobal().sValue("-CertificateDomainName" ,""), Path.GetExtension(lsLogPathFileBase)));

                return lsLogPathFileBase;
            }
        }

        public static bool bIsSetupCert(X509Certificate2 aoCertificate)
        {
            bool lbIsSetupCert = false;

            if ( null == Env.oDomainProfile && Env.sCertName(aoCertificate).ToLower() == tvProfile.oGlobal().sValue(Env.sStarCertNameKey, "").ToLower() )
                return false;

            if ( null == Env.oDomainProfile || Env.sCertName(aoCertificate).ToLower() != Env.oDomainProfile.sValue(Env.sStarCertNameKey, "").ToLower() )
            {
                if ( aoCertificate.Subject.Contains(Env.sNewClientSetupCertName) || aoCertificate.Subject.Contains("*") )
                    lbIsSetupCert = true;
                else
                    foreach (X509Extension loExt in aoCertificate.Extensions)
                    {
                        if ( "Subject Alternative Name" == loExt.Oid.FriendlyName )
                        {
                            AsnEncodedData  loAsnEncodedData = new AsnEncodedData(loExt.Oid, loExt.RawData);
                            string          lsSanList = loAsnEncodedData.Format(true);
                                            if ( lsSanList.Contains(Env.sNewClientSetupCertName) || lsSanList.Contains("*") )
                                            {
                                                lbIsSetupCert = true;
                                                break;
                                            }
                        }
                    }
            }

            return lbIsSetupCert;
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
                Environment.ExitCode = 1;
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
                    Environment.ExitCode = 1;
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

        public static bool bPfxToPem(string asAssemblyName, string asPfxPathFile, string asPfxPassword, string asPemPathFile, string asPemPassword)
        {
            bool lbPfxToPem = false;

            string  lsOpenSslFile = "openssl.zip";
            string  lsOpenSslPath = Path.GetFileNameWithoutExtension(lsOpenSslFile);
            string  lsOpenSslPathFile = tvProfile.oGlobal().sRelativeToExePathFile(lsOpenSslFile);
                    // Fetch OpenSsl module.
                    tvFetchResource.ToDisk(asAssemblyName
                            , String.Format("{0}{1}", Env.sFetchPrefix, lsOpenSslFile), lsOpenSslPathFile);
                    if ( !Directory.Exists(tvProfile.oGlobal().sRelativeToExePathFile(lsOpenSslPath)) )
                    {
                        string  lsZipDir = Path.GetDirectoryName(lsOpenSslPathFile);
                                ZipFile.ExtractToDirectory(lsOpenSslPathFile, String.IsNullOrEmpty(lsZipDir) ? "." : lsZipDir);
                    }
            string  lsDiscard = null;

            Directory.CreateDirectory(Path.GetDirectoryName(asPemPathFile));

            lbPfxToPem = Env.bRunPowerScript(out lsDiscard, null, tvProfile.oGlobal().sValue("-ScriptPfxToPem", @"
cd ""openssl""
C:PFXtoPEM2.cmd ""{PfxPathFile}"" ""{PemPathFile}"" -CertificateKey {PfxPassword} {PemPassword} 2>$null
                    ")
                    .Replace("{PfxPathFile}", asPfxPathFile)
                    .Replace("{PemPathFile}", asPemPathFile)
                    .Replace("{PfxPassword}", asPfxPassword)
                    .Replace("{PemPassword}", asPemPassword)
                    , false, true, false);

            return lbPfxToPem;
        }

        public static bool bReplaceInFiles(string asAssemblyName, tvProfile aoReplacementCmdLine)
        {
            return Env.bReplaceInFiles(asAssemblyName, aoReplacementCmdLine, false);
        }
        public static bool bReplaceInFiles(string asAssemblyName, tvProfile aoReplacementCmdLine, bool abSkipLog)
        {
            bool lbReplaceInFiles = false;

            string      lsBaseFile = "ReplaceText.exe";
            string      lsExePathFile = Path.Combine(tvProfile.oGlobal().sRelativeToExePathFile(Path.GetFileNameWithoutExtension(lsBaseFile)), lsBaseFile);
            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = lsExePathFile;
                        loProcess.StartInfo.Arguments = aoReplacementCmdLine.sCommandLine();
                        loProcess.StartInfo.UseShellExecute = false;
                        loProcess.StartInfo.RedirectStandardError = true;
                        loProcess.StartInfo.RedirectStandardInput = true;
                        loProcess.StartInfo.RedirectStandardOutput = true;
                        loProcess.StartInfo.CreateNoWindow = true;

                        // Fetch ReplaceText.exe.
                        Directory.CreateDirectory(Path.GetDirectoryName(lsExePathFile));
                        tvFetchResource.ToDisk(asAssemblyName
                                , String.Format("{0}{1}", Env.sFetchPrefix, lsBaseFile), lsExePathFile);
                        tvFetchResource.ToDisk(asAssemblyName
                                , String.Format("{0}{1}", Env.sFetchPrefix, lsBaseFile + ".txt"), lsExePathFile + ".txt");

            System.Windows.Forms.Application.DoEvents();

            Env.bPowerScriptSkipLog = abSkipLog;
            Env.bPowerScriptError = false;
            Env.sPowerScriptOutput = null;

            loProcess.Start();
            loProcess.BeginErrorReadLine();
            loProcess.BeginOutputReadLine();

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(tvProfile.oGlobal().dValue("-ReplaceTextTimeoutSecs", 600));

            if ( !abSkipLog )
                Env.LogIt(loProcess.StartInfo.Arguments);

            // Wait for the process to finish.
            while ( !Env.bMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !Env.bMainLoopStopped )
                    Thread.Sleep(tvProfile.oGlobal().iValue("-ReplaceTextSleepMS", 200));
            }

            if ( !Env.bMainLoopStopped )
            {
                if ( !loProcess.HasExited && !abSkipLog )
                    Env.LogIt(tvProfile.oGlobal().sValue("-ReplaceTextTimeoutMsg", "*** replacement sub-process timed-out ***\r\n\r\n"));

                int liExitCode = -1;
                    try { liExitCode = loProcess.ExitCode; } catch {}

                if ( Env.bPowerScriptError || liExitCode != 0 || !loProcess.HasExited )
                {
                    if ( !abSkipLog )
                        Env.LogIt(Env.sExceptionMessage("The replacement sub-process experienced a critical failure."));
                }
                else
                {
                    lbReplaceInFiles = true;

                    if ( !abSkipLog )
                        Env.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            Env.bKillProcess(loProcess);

            return lbReplaceInFiles;
        }

        public static bool bRunPowerScript(string asScript)
        {
            string lsDiscard = null;

            return Env.bRunPowerScript(out lsDiscard, asScript);
        }
        public static bool bRunPowerScript(out string asOutput, string asScript)
        {
            return Env.bRunPowerScript(out asOutput, null, asScript, false, false, false);
        }
        public static bool bRunPowerScript(
                  out string asOutput
                , string asSingleSessionScriptPathFile
                , string asScript
                , bool abOpenOrCloseSingleSession
                , bool abSkipLog
                , bool abRetryOutput
                )
        {
            bool            lbSingleSessionEnabled = tvProfile.oGlobal().bValue("-SingleSessionEnabled", false);
                            if ( !lbSingleSessionEnabled && abOpenOrCloseSingleSession )
                            {
                                // -SingleSessionEnabled is false, so ignore the single session "open" and "close" scripts.
                                asOutput = null;
                                return true;
                            }
            bool            lbRunPowerScript = false;
            string          lsScriptPathFile = tvProfile.oGlobal().sRelativeToExePathFile(
                                    tvProfile.oGlobal().sValue("-PowerScriptPathFile", "PowerScript.ps1")
                                    .Replace(".ps1", tvProfile.oGlobal().sValue("-CertificateDomainName" ,"") + ".ps1"));
            string          lsScript = null;
                            string          lsKludgeHide1 = Guid.NewGuid().ToString("N");   // Apparently Regex can't handle "++" in the text to process.
                            StringBuilder   lsbReplaceProfileTokens = new StringBuilder(asScript.Replace("++", lsKludgeHide1));

                            // Look for string tokens of the form: {ProfileKey}, where "ProfileKey" is
                            // only expected to be found in the profile if it is prefixed by a hyphen.
                            bool    lbUseLiteralsOnly = tvProfile.oGlobal().bUseLiteralsOnly;
                            Regex   loRegex = new Regex("{(.*?)}");
                                    tvProfile.oGlobal().bUseLiteralsOnly = true;
                                    foreach (Match loMatch in loRegex.Matches(lsbReplaceProfileTokens.ToString()))
                                    {
                                        string  lsKey = loMatch.Groups[1].ToString();
                                                if ( lsKey.StartsWith("-") && tvProfile.oGlobal().ContainsKey(lsKey) )
                                                    lsbReplaceProfileTokens.Replace(loMatch.Groups[0].ToString(), tvProfile.oGlobal()[lsKey].ToString());
                                    }
                                    tvProfile.oGlobal().bUseLiteralsOnly = lbUseLiteralsOnly;

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
                    if ( String.IsNullOrEmpty(asSingleSessionScriptPathFile) )
                        lsProcessArgs = String.Format(lsProcessArgs, lsScriptPathFile, "");
                    else
                        lsProcessArgs = String.Format(lsProcessArgs, asSingleSessionScriptPathFile, lsScriptPathFile);
            string  lsLogScript = !lsScript.Contains("-CertificateKey") ? lsScript : lsScript.Substring(0, lsScript.IndexOf("-CertificateKey"));

            if ( !abSkipLog && (!lbSingleSessionEnabled || !abOpenOrCloseSingleSession) )
                Env.LogIt(lsLogScript);

            Env.bPowerScriptSkipLog = abSkipLog;
            Env.sPowerScriptOutput = null;

            int liPowerScriptOutputRetries = tvProfile.oGlobal().iValue("-PowerScriptOutputRetries", 3);

            for (int i=0; i < liPowerScriptOutputRetries; i++)
            {
                Process loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = lsProcessPathFile;
                        loProcess.StartInfo.Arguments = lsProcessArgs;
                        loProcess.StartInfo.WorkingDirectory = Path.GetDirectoryName(lsScriptPathFile);
                        loProcess.StartInfo.UseShellExecute = false;
                        loProcess.StartInfo.RedirectStandardError = true;
                        loProcess.StartInfo.RedirectStandardInput = true;
                        loProcess.StartInfo.RedirectStandardOutput = true;
                        loProcess.StartInfo.CreateNoWindow = true;

                System.Windows.Forms.Application.DoEvents();

                Env.bPowerScriptError = false;

                try
                {
                    loProcess.Start();
                }
                catch (Exception ex)
                {
                    Env.LogIt(ex.Message);

                    if ( i <= liPowerScriptOutputRetries - 1 )
                        Env.LogIt(String.Format("Proceeding to retry #{0} ...", i + 1));
                }

                loProcess.BeginErrorReadLine();
                loProcess.BeginOutputReadLine();

                DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(tvProfile.oGlobal().dValue("-PowerScriptTimeoutSecs", 300));

                // Wait for the process to finish.
                while ( !Env.bMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
                {
                    System.Windows.Forms.Application.DoEvents();

                    if ( !Env.bMainLoopStopped )
                        Thread.Sleep(tvProfile.oGlobal().iValue("-PowerScriptSleepMS", 200));
                }

                if ( !Env.bMainLoopStopped )
                {
                    if ( !loProcess.HasExited )
                        Env.LogIt(tvProfile.oGlobal().sValue("-PowerScriptTimeoutMsg", "*** sub-process timed-out ***\r\n\r\n"));

                    int liExitCode = -1;
                        try { liExitCode = loProcess.ExitCode; } catch {}

                    if ( Env.bPowerScriptError || liExitCode != 0 || !loProcess.HasExited )
                    {
                        if ( !abSkipLog )
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

                if ( !abRetryOutput || !String.IsNullOrEmpty(Env.sPowerScriptOutput) )
                {
                    break;
                }
                else
                {
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(tvProfile.oGlobal().iValue("-PowerScriptOutputRetryMS", 1000));
                }
            }

            asOutput = Env.sPowerScriptOutput;

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

        public static void ProcessExitWait()
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(tvProfile.oGlobal().iValue("-ProcessExitWaitMS", 1000));
        }

        /// <summary>
        /// Returns the current IIS port 443 bound certificate (if any), otherwise whatever's in the local store (by name).
        /// </summary>
        public static X509Certificate2 oCurrentCertificate()
        {
            return Env.oCurrentCertificate(null);
        }
        public static X509Certificate2 oCurrentCertificate(string asCertName)
        {
            string[]        lsSanArray = null;
            ServerManager   loDiscard1 = null;
            Site            loDiscard2 = null;
            Site            loDiscard3 = null;

            return Env.oCurrentCertificate(asCertName, lsSanArray, out loDiscard1, out loDiscard2, out loDiscard3);
          }
        public static X509Certificate2 oCurrentCertificate(string asCertName, out ServerManager aoServerManager)
        {
            string[]        lsSanArray = null;
            Site            loDiscard1 = null;
            Site            loDiscard2 = null;

            return Env.oCurrentCertificate(asCertName, lsSanArray, out aoServerManager, out loDiscard1, out loDiscard2);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out Site aoSiteFound
                )
        {
            ServerManager   loDiscard1 = null;
            Site            loDiscard2 = null;

            return Env.oCurrentCertificate(asCertName, asSanArray, out loDiscard1, out aoSiteFound, out loDiscard2);
        }
        public static X509Certificate2 oCurrentCertificate(
                  string asCertNameOrSanItem
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                )
        {
            Site loDiscard = null;

            return Env.oCurrentCertificate(asCertNameOrSanItem, asSanArray, out aoServerManager, out aoSiteFound, out loDiscard);
        }
        public static X509Certificate2 oCurrentCertificate(
                  string asCertNameOrSanItem
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , out Site aoPrimarySiteForDefaults
                )
        {
            X509Certificate2  loCurrentCertificate = null;

            // ServerManager is the IIS ServerManager. It gives us the website object for binding the cert.
            aoServerManager = null;
            aoSiteFound = null;
            aoPrimarySiteForDefaults = null;

            try
            {
                string  lsCertNameOrSanItem = Env.sCertName(asCertNameOrSanItem);
                bool    lbLogCertificateStatus = tvProfile.oGlobal().bValue("-LogCertificateStatus", false);
                        if ( lbLogCertificateStatus )
                        {
                            Env.LogIt(new String('*', 80));
                            Env.LogIt("Env.oCurrentCertificate(asCertNameOrSanItem, asSanArray)");
                            Env.LogIt(String.Format("Env.oCurrentCertificate({0}, {1})"
                                        , null == asCertNameOrSanItem ? "null" : String.Format("\"{0}\"", lsCertNameOrSanItem)
                                        , null == asSanArray ? "null" : String.Format("\"{0}\"", String.Join(",", asSanArray))
                                        ));
                        }

                if (       null != Env.oGetCertServiceFactory
                        && null != Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate
                        && asCertNameOrSanItem == Env.sCertName(Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate)
                        && tvProfile.oGlobal().bValue("-CertificateSetupDone", false)
                        && null == asSanArray   // Given a non-null SAN array means do a full search and return the various objects found.
                        )
                {
                    loCurrentCertificate = Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate;

                    if ( lbLogCertificateStatus )
                        Env.LogIt(String.Format("Using certificate already found for the communications channel (\"{0}\"=\"{1}\").", Env.sCertName(loCurrentCertificate), loCurrentCertificate.Thumbprint));

                    if ( !tvProfile.oGlobal().bValue("-UseNonIISBindingOnly", false) )
                        aoServerManager = new ServerManager();
                }
                else
                {
                    lsCertNameOrSanItem = null == asCertNameOrSanItem ? null : lsCertNameOrSanItem.ToLower();

                    if ( lbLogCertificateStatus )
                        Env.LogIt("First, look thru existing IIS websites ...");

                    if ( !tvProfile.oGlobal().bValue("-UseNonIISBindingOnly", false) )
                    {
                        Binding loBindingFound = null;

                        aoServerManager = new ServerManager();

                        string  lsSiteName = null;
                                if ( lsCertNameOrSanItem != tvProfile.oGlobal().sValue(Env.sStarCertNameKey, "").ToLower() )
                                    lsSiteName = lsCertNameOrSanItem;
                                else
                                if ( null != asSanArray && 0 != asSanArray.Length )
                                    lsSiteName = asSanArray[0].ToLower();

                        foreach (Site loSite in aoServerManager.Sites)
                        {
                            if ( loSite.Name == lsSiteName )
                            {
                                if ( lbLogCertificateStatus )
                                    Env.LogIt(String.Format("Site found by matching certificate name or SAN item name (\"{0}\"). Now looking for bindings by matching certificate thumbprint ...", loSite.Name));

                                foreach (Binding loBinding in loSite.Bindings)
                                {
                                    if ( lbLogCertificateStatus )
                                        Env.LogIt(String.Format("Binding found (\"{0}\"). Now looking for a matching certificate - ie. a non-null \"Binding.CertificateHash={1}\" - (and a matching name) ..."
                                                , loBinding.BindingInformation, null == loBinding.CertificateHash ? "null" : Env.sCertThumbprintFromBindingHash(loBinding)));

                                    if ( null != loBinding.CertificateHash )
                                    foreach (X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                                    {
                                        if ( loCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash)
                                                && lsCertNameOrSanItem == Env.sCertName(loCertificate).ToLower() )
                                        {
                                            // Binding found that matches certificate thumbprint and name.
                                            loCurrentCertificate = loCertificate;
                                            aoSiteFound = loSite;
                                            loBindingFound = loBinding;

                                            if ( lbLogCertificateStatus )
                                                Env.LogIt(String.Format("'My' store certificate matches (\"{0}\" - by thumbprint & certificate name).", loCertificate.Thumbprint));
                                            break;
                                        }
                                        else
                                        if ( lbLogCertificateStatus )
                                            Env.LogIt(String.Format("'My' store certificate does NOT match (\"{0}\"=\"{1}\").", Env.sCertName(loCertificate), loCertificate.Thumbprint));
                                    }

                                    if ( null != loBindingFound )
                                        break;
                                }

                                if ( null == loBindingFound && lbLogCertificateStatus )
                                    Env.LogIt("No match found. No binding selected.");
                            }

                            if ( null != aoSiteFound )
                                break;
                        }

                        if ( null == aoSiteFound )
                        foreach (Site loSite in aoServerManager.Sites)
                        {
                            if ( lbLogCertificateStatus )
                                Env.LogIt(String.Format("Site found (\"{0}\") in sequence - not by matching name. Now looking for bindings by matching certificate thumbprint ...", loSite.Name));

                            foreach (Binding loBinding in loSite.Bindings)
                            {
                                if ( lbLogCertificateStatus )
                                    Env.LogIt(String.Format("Binding found (\"{0}\"). Now looking for a matching certificate - ie. a non-null \"Binding.CertificateHash={1}\" - (and a matching name - if asCertNameOrSanItem is not null) ..."
                                            , loBinding.BindingInformation, null == loBinding.CertificateHash ? "null" : Env.sCertThumbprintFromBindingHash(loBinding)));

                                if ( null != loBinding.CertificateHash )
                                foreach (X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                                {
                                    if ( loCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash)
                                        && (String.IsNullOrEmpty(lsCertNameOrSanItem)
                                            || lsCertNameOrSanItem == Env.sCertName(loCertificate).ToLower()) )
                                    {
                                        // Binding found that matches certificate thumbprint and name (or no name).
                                        loCurrentCertificate = loCertificate;
                                        aoSiteFound = loSite;
                                        loBindingFound = loBinding;

                                        if ( lbLogCertificateStatus )
                                            Env.LogIt(String.Format("'My' store certificate matches (\"{0}\" - {1}).", loCertificate.Thumbprint
                                                    , lsCertNameOrSanItem == Env.sCertName(loCertificate).ToLower() ? "by thumbprint & certificate name" : "by thumbprint only, no name given"));
                                        break;
                                    }
                                    else
                                    if ( lbLogCertificateStatus )
                                        Env.LogIt(String.Format("'My' store certificate does NOT match (\"{0}\"=\"{1}\").", Env.sCertName(loCertificate), loCertificate.Thumbprint));
                                }

                                if ( null != loBindingFound )
                                    break;
                            }

                            if ( null == loBindingFound && lbLogCertificateStatus )
                                Env.LogIt("No match found. No binding selected.");

                            if ( null == loBindingFound && null != Env.oChannelCertificate && Env.bIsSetupCert(Env.oChannelCertificate) )
                            {
                                if ( lbLogCertificateStatus )
                                    Env.LogIt(String.Format("Same site (\"{0}\"). Next, looking for a binding matching the setup certificate thumbprint (name match not required) ...", loSite.Name));

                                foreach (Binding loBinding in loSite.Bindings)
                                {
                                    if ( null != loBinding.CertificateHash && Env.oChannelCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash) )
                                    {
                                        // Binding found that matches the setup certificate.
                                        // loCurrentCertificate MUST be null in this case to avoid attempted removal of the "old certificate."
                                        loCurrentCertificate = null;
                                        aoSiteFound = loSite;
                                        loBindingFound = loBinding;

                                        if ( lbLogCertificateStatus )
                                            Env.LogIt(String.Format("Binding (\"{0}\") matches setup certificate thumbprint (\"{1}\")."
                                                                    , loBinding.BindingInformation, Env.oChannelCertificate.Thumbprint));
                                        break;
                                    }
                                    else
                                    if ( lbLogCertificateStatus )
                                        Env.LogIt(String.Format("Binding (\"{0}\") does NOT match setup certificate thumbprint (\"{1}\")."
                                                                , loBinding.BindingInformation, Env.oChannelCertificate.Thumbprint));
                                }

                                if ( null == loBindingFound && lbLogCertificateStatus )
                                    Env.LogIt("No match found. No binding selected.");
                            }

                            // Next, if no binding certificate match exists, look by binding hostname (ie. SNI)
                            // (note: this is just to populate aoSiteFound and loBindingFound).
                            if ( null == loBindingFound )
                            foreach (Binding loBinding in loSite.Bindings)
                            {
                                if ( null == loBinding.CertificateHash && !String.IsNullOrEmpty(lsCertNameOrSanItem)
                                        && lsSiteName == loBinding.Host.ToLower() && "https" == loBinding.Protocol )
                                {
                                    // Site found with a binding hostname that matches lsSiteName.
                                    aoSiteFound = loSite;
                                    loBindingFound = loBinding;
                                    break;
                                }

                                if ( null != loBindingFound )
                                    break;
                            }

                            // Finally, with no binding found, try to match against the site name.
                            // (note: this is just to populate aoSiteFound).
                            if ( null == loBindingFound && loSite.Name.ToLower() == lsSiteName )
                                aoSiteFound = loSite;

                            if ( null != aoSiteFound )
                                break;
                        }

                        // Finally, finally (no really), with no site found, find a related primary site to use for defaults.
                        string lsPrimaryCertName = null == asSanArray || 0 == asSanArray.Length ? null : asSanArray[0].ToLower();

                        if ( null == aoSiteFound && !String.IsNullOrEmpty(lsPrimaryCertName)
                                && lsCertNameOrSanItem != lsPrimaryCertName && "" == tvProfile.oGlobal().sValue(Env.sStarCertNameKey, "") )
                        {
                            // Use the primary site in the SAN array as defaults for any new site (to be created).
                            Env.oCurrentCertificate(lsPrimaryCertName, asSanArray, out aoPrimarySiteForDefaults);
                        }
                    }

                    if ( lbLogCertificateStatus && null == aoServerManager )
                        Env.LogIt("    IIS is not used by this instance.");

                    // As a last resort (so much for finally), use the oldest certificate found in the local 'My' store (by name).
                    if ( null == asSanArray && null == loCurrentCertificate && !String.IsNullOrEmpty(lsCertNameOrSanItem) )
                    {
                        if ( lbLogCertificateStatus )
                            Env.LogIt("Finally, looking for a certificate in the local 'My' store by name ...");
                
                        foreach (X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                        {
                            if ( lsCertNameOrSanItem == Env.sCertName(loCertificate).ToLower()
                                    && (null == loCurrentCertificate || loCertificate.NotAfter < loCurrentCertificate.NotAfter) )
                            {
                                loCurrentCertificate = loCertificate;

                                if ( lbLogCertificateStatus )
                                    Env.LogIt(String.Format("Certificate (by name) matches (\"{0}\"=\"{1}\" - looking for oldest).", Env.sCertName(loCertificate), loCertificate.Thumbprint));
                            }
                        }
                    }
                }

                if ( lbLogCertificateStatus )
                {
                    if ( null == asSanArray && null == loCurrentCertificate )
                        Env.LogIt("No current certificate found.");

                    Env.LogIt(new String('*', 80));
                }
            }
            catch (Exception ex)
            {
                Environment.ExitCode = 1;
                Env.LogIt(ex.Message);
                Env.LogIt("IIS may not be installed or fully configured. This will prevent certificate-to-site binding in IIS.");
            }

            return loCurrentCertificate;
        }

        public static string sExceptionMessage(Exception ex)
        {
            return String.Format("ServiceFault({0}): {1}{2}{3}{4}", ex.GetType().FullName, ex.Message, (null == ex.InnerException ? "": "; " + ex.InnerException.Message), Environment.NewLine, ex.StackTrace);
        }
        public static string sExceptionMessage(string asMessage)
        {
            return "ServiceFault: " + asMessage;
        }

        public static string sHostExePathFile()
        {
            Process loDiscardProcess = null;
            return Env.sHostExePathFile(out loDiscardProcess, false);
        }
        public static string sHostExePathFile(out Process aoProcessFound)
        {
            return Env.sHostExePathFile(out aoProcessFound, false);
        }
        public static string sHostExePathFile(out Process aoProcessFound, bool abQuietMode)
        {
            string  lsHostExePathFile = Env.sProcessExePathFile(Env.sHostProcess, out aoProcessFound);

            if ( String.IsNullOrEmpty(lsHostExePathFile) )
            {
                if ( !abQuietMode )
                    Env.LogIt("Host image can't be located on disk based on the currently running process. Trying typical locations ...");

                lsHostExePathFile = Path.Combine(Path.Combine(@"C:\ProgramData", Path.GetFileNameWithoutExtension(Env.sHostProcess)), Env.sHostProcess);
                if ( !File.Exists(lsHostExePathFile) )
                    lsHostExePathFile = Path.Combine(Path.Combine(@"C:\Program Files", Path.GetFileNameWithoutExtension(Env.sHostProcess)), Env.sHostProcess);

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

        public static string sHostLogText()
        {
            return Env.sHostLogText(false);
        }
        public static string sHostLogText(bool abUnfiltered)
        {
            if ( tvProfile.oGlobal().sLoadedPathFile != tvProfile.oGlobal().sDefaultPathFile )
                return null;

            string          lsHostLogText1 = null;
            string          lsHostLogText2 = null;
            string          lsHostLogSpec = tvProfile.oGlobal().sValue("-HostLogSpec", String.Format(@"Logs\{0}Log*", Env.sHostProcess));
            string          lsHostLogPathFiles = null;
                            if ( "" == Path.GetPathRoot(lsHostLogSpec) )
                            {
                                lsHostLogPathFiles = Path.Combine(Path.GetDirectoryName(Env.sHostExePathFile()), lsHostLogSpec);
                            }
                            else
                            {
                                lsHostLogPathFiles = lsHostLogSpec;
                            }
            string          lsHostLogPath = Path.GetDirectoryName(lsHostLogPathFiles);
            string          lsHostLogFiles = Path.GetFileName(lsHostLogPathFiles);
            IOrderedEnumerable<FileSystemInfo> loFileSysInfoList = null;
                            if ( Directory.Exists(lsHostLogPath) )
                                loFileSysInfoList = new DirectoryInfo(lsHostLogPath).GetFileSystemInfos(lsHostLogFiles).OrderByDescending(a => a.LastWriteTime);
            string          lsHostLogPathFile1 = null;
                            if ( null != loFileSysInfoList && loFileSysInfoList.Count() > 0 ) 
                                lsHostLogPathFile1 = loFileSysInfoList.First().FullName;
                            else
                                lsHostLogPathFile1 = String.Format("Error locating today's \"{0}\" file.", lsHostLogPathFiles);
            string          lsHostLogPathFile2 = null;
                            if ( null != loFileSysInfoList && loFileSysInfoList.Count() > 1 ) 
                                lsHostLogPathFile2 = loFileSysInfoList.Skip(1).First().FullName;
                            else
                                lsHostLogPathFile2 = String.Format("Error locating yesterday's \"{0}\" file.", lsHostLogPathFiles);
            string          lsDelimiterBeg = String.Format("{0}  BEGIN   {{0}}   BEGIN  {0}", new String('*', 26));
            string          lsDelimiterEnd = String.Format("{0}  END     {{0}}     END  {0}", new String('*', 26));
            string          lsFinalFormat = "{0}{1}{0}{2}{3}{0}";

            if ( abUnfiltered )
            {
                if ( File.Exists(lsHostLogPathFile1) )
                    lsHostLogText1 = File.ReadAllText(lsHostLogPathFile1);

                if ( File.Exists(lsHostLogPathFile2) )
                    lsHostLogText2 = File.ReadAllText(lsHostLogPathFile2);
            }
            else
            {
                string[]        lsHostLogTextArray = null;
                StringBuilder   lsbHostLogTextFiltered = null;

                if ( File.Exists(lsHostLogPathFile1) )
                {
                    lsHostLogTextArray = File.ReadAllText(lsHostLogPathFile1).Split(Environment.NewLine.ToCharArray());
                    lsbHostLogTextFiltered = new StringBuilder();

                    foreach(string lsLine in lsHostLogTextArray)
                        if ( lsLine.Contains(tvProfile.oGlobal().sValue("-HostLogFilter", "-CommandEXE=")) )
                            lsbHostLogTextFiltered.AppendLine(lsLine);

                    lsHostLogText1 = lsbHostLogTextFiltered.ToString();
                }

                if ( File.Exists(lsHostLogPathFile2) )
                {
                    lsHostLogTextArray = File.ReadAllText(lsHostLogPathFile2).Split(Environment.NewLine.ToCharArray());
                    lsbHostLogTextFiltered = new StringBuilder();

                    foreach(string lsLine in lsHostLogTextArray)
                        if ( lsLine.Contains(tvProfile.oGlobal().sValue("-HostLogFilter", "-CommandEXE=")) )
                            lsbHostLogTextFiltered.AppendLine(lsLine);

                    lsHostLogText2 = lsbHostLogTextFiltered.ToString();
                }
            }
        

            return    String.Format(lsFinalFormat, Environment.NewLine, String.Format(lsDelimiterBeg, lsHostLogPathFile1), lsHostLogText1, String.Format(lsDelimiterEnd, lsHostLogPathFile1))
                    + String.Format(lsFinalFormat, Environment.NewLine, String.Format(lsDelimiterBeg, lsHostLogPathFile2), lsHostLogText2, String.Format(lsDelimiterEnd, lsHostLogPathFile2));
        }

        public static string sHostProfile()
        {
            return Env.sHostProfile(false);
        }
        public static string sHostProfile(bool abUnfiltered)
        {
            string lsHostProfile = null;

            if ( abUnfiltered )
            {
                lsHostProfile = File.ReadAllText(Env.sHostExePathFile() + ".txt.backup.txt");
            }
            else
            {
                tvProfile   loHostProfile = new tvProfile(Env.sHostExePathFile() + ".txt.backup.txt", false);
                            loHostProfile.bEnableFileLock = false;
                tvProfile   loHostProfileFiltered = loHostProfile.oOneKeyProfile("-AddTasks");
                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-BackupFiles"))
                                loHostProfileFiltered.Add(loEntry);
                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-BackupSet"))
                                loHostProfileFiltered.Add(loEntry);
                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-CleanupFiles"))
                                loHostProfileFiltered.Add(loEntry);
                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-CleanupSet"))
                                loHostProfileFiltered.Add(loEntry);

                lsHostProfile = loHostProfileFiltered.ToString();
            }

            return lsHostProfile;
        }

        public static string sProcessExePathFile(string asExePathFile)
        {
            Process loProcessFound = null;

            return Env.sProcessExePathFile(asExePathFile, out loProcessFound);
        }
        public static string sProcessExePathFile(string asExePathFile, out Process aoProcessFound)
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


        private static string sCertThumbprintFromBindingHash(Binding aoBinding)
        {
            StringBuilder   lsbCertThumbprintFromBindingHash = new StringBuilder();
                            for (int i = 0; i < aoBinding.CertificateHash.Length; i++)
                                lsbCertThumbprintFromBindingHash.AppendFormat("{0:X2}", aoBinding.CertificateHash[i]);

            return lsbCertThumbprintFromBindingHash.ToString();
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
                lbDone = abAppendOutput || !File.Exists(lsOutputPathFile);

                // If the given pathfile already exists, create a variation on the dated
                // filename by appending an integer (see liUniqueFilenameSuffix above).
                // Keep trying until a unique dated pathfile is identified.
                if ( !lbDone )
                    lsOutputFilename = lsBaseFilename + "." + (++liUniqueFilenameSuffix).ToString() + lsBaseFileExt;
            }
            while ( !lbDone );

            return lsOutputPathFile;
        }

        public static bool      bMainLoopStopped = false;
        public static bool      bPowerScriptError = false;
        public static bool      bPowerScriptSkipLog = false;
        public static int       iNonErrorBounceSecs = 0;
        public static string    sPowerScriptOutput = null;
        public static string    sFetchPrefix = "Resources.Fetch.";
        public static string    sHostProcess = "GoPcBackup.exe";
        public static string    sNewClientSetupPfxName = "GgcSetup.pfx";
        public static string    sNewClientSetupCertName = "GetCertClientSetup";
        public static string    sStarCertNameKey = "-StarCertName";
        public static string    sWcfLogFile  = "WcfLog.txt";         // Must be changed in the WCF config too.

        /// <summary>
        /// Write the given asText to a text file as well as to
        /// the output console of the UI window (if it exists).
        /// </summary>
        /// <param name="asText">The text string to log.</param>
        public static void LogIt(string asText)
        {
            Env.LogIt(asText, false);
        }
        public static void LogIt(string asText, bool abBaseIsParentFolder)
        {
            StreamWriter    loStreamWriter = null;
            string          lsLogPathFile = Env.sLogPathFile;
                            if ( abBaseIsParentFolder )
                            {
                                string  lsParentPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(lsLogPathFile)));
                                string  lsLogBasePath = Path.GetDirectoryName(lsLogPathFile).Replace(Path.GetDirectoryName(Path.GetDirectoryName(lsLogPathFile)), "");

                                lsLogPathFile = Path.Combine(lsParentPath + lsLogBasePath, Path.GetFileName(lsLogPathFile));
                            }
            string          lsPath = Path.GetDirectoryName(lsLogPathFile);
                            if ( !Directory.Exists(lsPath) )
                                Directory.CreateDirectory(lsPath);
            string          lsFormat = tvProfile.oGlobal().sValueNoTrim(
                                    "-LogEntryDateTimeFormatPrefix", "yyyy-MM-dd hh:mm:ss:fff tt  ");

            try
            {
                loStreamWriter = new StreamWriter(lsLogPathFile, true);
                loStreamWriter.WriteLine(DateTime.Now.ToString(lsFormat) + asText);
            }
            catch
            {
                try
                {
                    // Give it one more try ...
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(200);

                    loStreamWriter = new StreamWriter(lsLogPathFile, true);
                    loStreamWriter.WriteLine(DateTime.Now.ToString(lsFormat) + asText);
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

        public static void ScheduleBounce(int aiSecondsUntilBounce, bool abSkipLog)
        {
            if (       !Env.oDomainProfile.bValue("-EnableRetryBounce", false)
                    || DateTime.Now < Env.oDomainProfile.dtValue("-MaintenanceWindowBeginTime", DateTime.MaxValue)
                    || DateTime.Now > Env.oDomainProfile.dtValue("-MaintenanceWindowEndTime", DateTime.MinValue)
                            .AddMinutes(tvProfile.oGlobal().iValue("-AddMinsToMaintenanceWindow", 0))
                            .AddSeconds(-aiSecondsUntilBounce)  // Subtract the seconds until bounce.
                    )
                return;

            string lsDiscard = null;

            Env.bRunPowerScript(out lsDiscard, null
                    , tvProfile.oGlobal().sValue("-ScriptBounce", "shutdown.exe /r /t {SecondsUntilBounce} /c \"{ExeName} initiated bounce for '{CertificateDomainName}' domain\"")
                            .Replace("{SecondsUntilBounce}", aiSecondsUntilBounce.ToString())
                            .Replace("{ExeName}", Path.GetFileNameWithoutExtension(tvProfile.oGlobal().sExePathFile))
                            .Replace("{CertificateDomainName}", tvProfile.oGlobal().sValue("-CertificateDomainName" ,""))
                    , false, abSkipLog, false);
        }

        public static void ScheduleBounceOnErrorCancel(bool abSkipLog)
        {
            if ( !tvProfile.oGlobal().bValue("-Auto", false) || tvProfile.oGlobal().bValue("-Setup", false) )
                return;

            File.Delete(Env.sBounceOnErrorLockPathFile);

            string  lsDiscard = null;
                    Env.bRunPowerScript(out lsDiscard, null
                            , tvProfile.oGlobal().sValue("-ScriptBounceCancel", "shutdown.exe /a"), false, abSkipLog, false);
        }

        public static void ScheduleBounceReset()
        {
            if ( !tvProfile.oGlobal().bValue("-Auto", false) || tvProfile.oGlobal().bValue("-Setup", false) )
                return;

            // Bounce locks can only be reset on subsequent days.

            if ( DateTime.Now.Date != new FileInfo(Env.sBounceOnErrorLockPathFile).LastWriteTime.Date )
                File.Delete(Env.sBounceOnErrorLockPathFile);

            if ( DateTime.Now.Date != new FileInfo(Env.sBounceOnNonErrorLockPathFile).LastWriteTime.Date )
                File.Delete(Env.sBounceOnNonErrorLockPathFile);
        }

        public static void ScheduleNonErrorBounce(tvProfile aoDomainProfile)
        {
            Env.ScheduleNonErrorBounce(false);
        }
        public static void ScheduleNonErrorBounce(bool abSkipLog)
        {
            if ( !tvProfile.oGlobal().bValue("-Auto", false) || tvProfile.oGlobal().bValue("-Setup", false) || !tvProfile.oGlobal().bValue("-EnableRetryBounce", true) )
                return;

            if ( !File.Exists(Env.sBounceOnNonErrorLockPathFile) )
            {
                File.WriteAllText(Env.sBounceOnNonErrorLockPathFile, tvProfile.oGlobal().sValue("-CertificateDomainName" ,""));

                Env.ScheduleBounce(Env.iNonErrorBounceSecs, abSkipLog);
            }
        }

        public static void ScheduleOnErrorBounce()
        {
            Env.ScheduleOnErrorBounce(false);
        }
        public static void ScheduleOnErrorBounce(bool abDuringStartup)
        {
            if ( !tvProfile.oGlobal().bValue("-Auto", false) || tvProfile.oGlobal().bValue("-Setup", false) || !tvProfile.oGlobal().bValue("-EnableRetryBounce", true) )
                return;

            if ( !File.Exists(Env.sBounceOnErrorLockPathFile) )
            {
                File.WriteAllText(Env.sBounceOnErrorLockPathFile, tvProfile.oGlobal().sValue("-CertificateDomainName" ,""));

                Env.ScheduleBounce(tvProfile.oGlobal().iValue("-OnErrorBounceSecs", 600), abDuringStartup);
            }
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

            bool lbLogCertificateStatus = tvProfile.oGlobal().bValue("-LogCertificateStatus", false);

            if ( lbLogCertificateStatus )
            {
                Env.LogIt(new String('*', 51));

                if ( null == Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate )
                    Env.LogIt("Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate=null");
                else
                    Env.LogIt(String.Format("Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate=(\"{0}\"=\"{1}\")"
                            , Env.sCertName(Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate), Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate.Thumbprint));
            }

            if ( null == Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate
                    || (loProfile.bValue("-CertificateSetupDone", false) && Env.bIsSetupCert(Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate)) )
            {
                if ( null == Env.oChannelCertificate )
                {
                    try
                    {
                        if ( "" != loProfile.sValue("-ContactEmailAddress" ,"") )
                            Env.oChannelCertificate = Env.oCurrentCertificate(loProfile.sValue("-CertificateDomainName" ,""));

                        if ( null == Env.oChannelCertificate && !loProfile.bValue("-CertificateSetupDone", false)
                                && (!tvProfile.oGlobal().bValue("-Auto", false) || tvProfile.oGlobal().bValue("-Setup", false))
                                && (!loProfile.bValue("-NoPrompts", false) || (!loProfile.bValue("-LicenseAccepted", false) && !loProfile.bValue("-AllConfigWizardStepsCompleted", false))) )
                        {
                            System.Windows.Forms.OpenFileDialog loOpenDialog = new System.Windows.Forms.OpenFileDialog();
                                                                loOpenDialog.FileName = Env.sNewClientSetupPfxName;
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
                                        Env.oChannelCertificate = new X509Certificate2(loOpenDialog.FileName, lsPassword);

                                        if ( null != Env.oChannelCertificate )
                                            loProfile["-NoPrompts"] = false;
                                    }
                                    catch (Exception ex)
                                    {
                                        Environment.ExitCode = 1;
                                        loProfile.bExit = true;

                                        Env.LogIt(Env.sExceptionMessage(ex));
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Environment.ExitCode = 1;
                        loProfile.bExit = true;

                        Env.LogIt(Env.sExceptionMessage(ex));

                        throw ex;
                    }
                }

                if ( null != Env.oChannelCertificate )
                    Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate = Env.oChannelCertificate;
            }

            if ( !lbLogCertificateStatus )
            {
                if ( null == Env.oChannelCertificate )
                    Env.oChannelCertificate = Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate;
            }
            else
            {
                if ( null == Env.oChannelCertificate )
                {
                    Env.LogIt("Env.oChannelCertificate=null");

                    if ( !loProfile.bValue("-CertificateSetupDone", false) )
                        Env.LogIt("    Note: -CertificateSetupDone=False (set this value True if you know the current domain certificate already supersedes the setup certificate.");

                    Env.oChannelCertificate = Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate;

                    Env.LogIt("Env.oChannelCertificate=Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate");
                }

                if ( null == Env.oChannelCertificate )
                    Env.LogIt("Env.oChannelCertificate=null");
                else
                    Env.LogIt(String.Format("Env.oChannelCertificate=(\"{0}\"=\"{1}\")"
                            , Env.sCertName(Env.oChannelCertificate), Env.oChannelCertificate.Thumbprint));

                if ( null == Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate )
                    Env.LogIt("Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate=null");
                else
                    Env.LogIt(String.Format("Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate=(\"{0}\"=\"{1}\")"
                            , Env.sCertName(Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate), Env.oGetCertServiceFactory.Credentials.ClientCertificate.Certificate.Thumbprint));

                Env.LogIt(new String('*', 51));
            }
        }
    }
}
