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
    public partial class Env
    {
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
        /// Returns the current server computer name.
        /// </summary>
        public static string sComputerName
        {
            get
            {
                return Env.oEnvProfile.sValue("-COMPUTERNAME", "Computer name not found.");
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

        public static void AddArgs(tvProfile aoResults, params object[] aoArgs)
        {
            if ( null == aoArgs )
                return;

            for ( int i=0; i < aoArgs.Length; i++ )
                aoResults.Add("-Arg" + (i + 1).ToString(), aoArgs[i]);
        }
    }
}
