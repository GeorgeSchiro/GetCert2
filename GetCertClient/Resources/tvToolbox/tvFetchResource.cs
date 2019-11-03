using System;
using System.Windows;
using System.Windows.Input;
using System.IO;
using System.Reflection;


namespace tvToolbox
{
    /// <summary>
    /// This utility class supports the embedding of DLLs in an EXE that uses them.
    /// This approach allows for a primary EXE to serve as its own setup program.
    ///
    /// This functionality can also be used to package any other files with
    /// the primary EXE (eg. DLLs, configuration files as well as other EXEs
    /// and their support files).
    /// </summary>
    internal class tvFetchResource
    {
        private tvFetchResource(){}

        /// <summary>
        /// Fetches an embedded resource from the currently executing assembly. It is
        /// written to disk to the same folder that contains the assembly. The namespace
        /// of the first class found in the assembly is used by default. If this fails,
        /// try including the namespace as the first argument.
        /// </summary>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        internal static void ToDisk(string asResourceName)
        {
            btArray(null, asResourceName, true, null);
        }

        /// <summary>
        /// Fetches an embedded resource from the currently executing assembly. The namespace
        /// of the first class found in the assembly is used by default. If this fails,
        /// try including the namespace as the first argument.
        ///
        /// The resource is written to disk at asPathFile.
        /// </summary>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        /// <param name="asPathFile">
        /// The location where the embedded resource is written to disk. If this parameter is null,
        /// the given resource name with the location of the currently running EXE is used instead.
        /// </param>
        internal static void ToDisk(string asResourceName, string asPathFile)
        {
            btArray(null, asResourceName, true, asPathFile);
        }

        /// <summary>
        /// Fetches an embedded resource from the currently executing assembly.
        ///
        /// The resource is written to disk at asPathFile.
        /// </summary>
        /// <param name="asNamespace">The namespace of the embedded resource to fetch.</param>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        /// <param name="asPathFile">
        /// The location where the embedded resource is written to disk. If this parameter is null,
        /// the given resource name with the location of the currently running EXE is used instead.
        /// </param>
        internal static void ToDisk(string asNamespace, string asResourceName, string asPathFile)
        {
            btArray(asNamespace, asResourceName, true, asPathFile);
        }

        /// <summary>
        /// Fetches a byte array (ie. an embedded resource) from the currently executing assembly.
        /// The namespace of the first class found in the assembly is used by default. If this fails,
        /// try including the namespace as the first argument.
        ///
        /// The byte array is written to disk at asPathFile.
        ///
        /// If the file already exists on disk (ie. asPathFile), a byte array is returned from the file.
        /// </summary>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        /// <param name="asPathFile">
        /// The location where the embedded resource is written to disk. If this parameter is null,
        /// the given resource name with the location of the currently running EXE is used instead.
        /// </param>
        /// <returns>
        /// A byte array that contains the resource fetched from the assembly.
        /// </returns>
        internal static byte[] btArrayToDisk(string asResourceName, string asPathFile)
        {
            return btArrayToDisk(null, asResourceName, asPathFile);
        }

        /// <summary>
        /// Fetches a byte array (ie. an embedded resource) from the currently executing assembly.
        ///
        /// The byte array is written to disk at asPathFile.
        ///
        /// If the file already exists on disk (ie. asPathFile), a byte array is returned from the file.
        /// </summary>
        /// <param name="asNamespace">The namespace of the embedded resource to fetch.</param>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        /// <param name="asPathFile">
        /// The location where the embedded resource is written to disk. If this parameter is null,
        /// the given resource name with the location of the currently running EXE is used instead.
        /// </param>
        /// <returns>
        /// A byte array that contains the resource fetched from the assembly.
        /// </returns>
        internal static byte[] btArrayToDisk(string asNamespace, string asResourceName, string asPathFile)
        {
            byte[] lbtArray = btArray(asNamespace, asResourceName, true, asPathFile);

            if ( null == lbtArray )
            {
                FileStream  loFileStream = null;

                try
                {
                    loFileStream = new FileStream(asPathFile, FileMode.Open, FileAccess.Read);
                    lbtArray = new Byte[loFileStream.Length];
                    loFileStream.Read(lbtArray, 0, (int)loFileStream.Length);
                }
                finally
                {
                    if ( null != loFileStream )
                        loFileStream.Close();
                }
            }

            return lbtArray;
        }

        /// <summary>
        /// Fetches a byte array (ie. an embedded resource) from the currently executing assembly.
        ///
        /// If abFetchToDisk is true, the byte array is written to disk at asPathFile.
        /// null is returned if the resource file already exists on disk.
        ///
        /// If abFetchToDisk is false, the contents of asResourceName is returned from the executing
        /// assembly without regard to what may already be on disk.
        /// </summary>
        /// <param name="asNamespace">The namespace of the embedded resource to fetch.</param>
        /// <param name="asResourceName">The name of the embedded resource to fetch.</param>
        /// <param name="abFetchToDisk">Fetch to disk if this boolean is true.</param>
        /// <param name="asPathFile">
        /// The location where the embedded resource is written to disk. If this parameter is null,
        /// the given resource name with the location of the currently running EXE is used instead.
        /// </param>
        /// <returns>
        /// A byte array that contains the resource fetched from the assembly.
        /// </returns>
        internal static byte[] btArray(string asNamespace, string asResourceName, bool abFetchToDisk, string asPathFile)
        {
            if ( null == asPathFile )
                asPathFile = Path.Combine(Path.GetDirectoryName(
                                Assembly.GetExecutingAssembly().Location), asResourceName);

            if ( File.Exists(asPathFile) && abFetchToDisk )
                return null;

            Stream loStream = null;
            FileStream loFileStream = null;
            Byte[] lbtArray = null;

            try
            {
                string lsResourceName = ( null != asNamespace && "" != asNamespace
                        ? asNamespace + "." + asResourceName
                        : Assembly.GetExecutingAssembly().GetName().Name + "." + asResourceName
                        );

                loStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(lsResourceName);
                if ( null == loStream )
                {
                    tvFetchResource.ErrorMessage(null, String.Format(
                              "The embedded resource ({0}) could not be found in the running assembly ({1})."
                            + Environment.NewLine + Environment.NewLine
                            + "Try specifying the namespace as the first argument."
                            , asResourceName, Assembly.GetExecutingAssembly().FullName));
                }
                else
                {
                    lbtArray = new Byte[loStream.Length];
                    loStream.Read(lbtArray, 0, (int)loStream.Length);

                    if ( abFetchToDisk )
                    {
                        loFileStream = new FileStream(asPathFile, FileMode.OpenOrCreate);
                        loFileStream.Write(lbtArray, 0, (int)loStream.Length);
                    }
                }
            }
            finally
            {
                if ( null != loStream)
                    loStream.Close();

                if ( null != loFileStream )
                    loFileStream.Close();
            }

            return lbtArray;
        }

        internal static void ErrorMessage(Window aoWindow, string asMessage)
        {
            Type lttvMessageBox = Type.GetType("tvMessageBox");
            if ( null != lttvMessageBox )
            {
                object  loErrMsg = Activator.CreateInstance(lttvMessageBox);

                lttvMessageBox.InvokeMember("ShowError", BindingFlags.InvokeMethod, null, loErrMsg
                                            , new object[]{aoWindow, asMessage, System.Windows.Forms.Application.ProductName});
            }
        }

        internal static void NetworkSecurityStartupErrorMessage()
        {
            ErrorMessage(null, "Network Security Exception: apparently you can't run this from the network."
                    + Environment.NewLine + Environment.NewLine
                    + "Try copying the software to a local PC before running it.");
        }

        internal static void ModelessMessageBox(tvProfile aoProfile, string asNamespace, string asTitle, string asMessage)
        {
            String lcsMsgBoxExeName = "MessageBox.exe";

            System.Diagnostics.Process[] loProcessArray = System.Diagnostics.Process.GetProcessesByName(
                    System.IO.Path.GetFileNameWithoutExtension(lcsMsgBoxExeName));

            if ( loProcessArray.Length < aoProfile.iValue("-ModelessMessageBoxMaxCount", 3) )
            {
                string lsMsgExePathFile = Path.Combine(
                        Path.GetDirectoryName(aoProfile.sExePathFile), lcsMsgBoxExeName);

                tvFetchResource.ToDisk(asNamespace, lcsMsgBoxExeName, lsMsgExePathFile);

                if ( File.Exists(lsMsgExePathFile) )
                    System.Diagnostics.Process.Start(lsMsgExePathFile, String.Format(
                              "-Title=\"{0}: {1}\" -Message=\"{2}\""
                            , asNamespace.Replace("\"", "'")
                            , asTitle.Replace("\"", "'")
                            , asMessage.Replace("\"", "'")
                            ));
            }
        }
    }
}
