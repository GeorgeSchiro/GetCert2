using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;


namespace tvToolbox
{
    /// <summary>
    /// Fetches a resource file by name (asResourceName) from the current executable.
    /// </summary>
    public delegate void FetchResourceToDisk(string asResourceName);

    /// <summary>
    /// Fetches a resource file by name (asResourceName) from the current executable
    /// using asNamespace. The resulting file is written to asPathFile.
    /// </summary>
    public delegate void FetchResourceToDisk2(string asNamespace, string asResourceName, string asPathFile);

    /// <summary>
    /// Default profile file actions specify how defaults are handled at runtime.
    /// </summary>
    public enum tvProfileDefaultFileActions
    {
        /// <summary>
        /// Automatically load the default profile file during application
        /// startup. Also, save the default profile file without prompts
        /// whenever new keys are automatically added (ie. whenever new
        /// keys with default values are referenced in the application
        /// code for the first time).
        /// </summary>
        AutoLoadSaveDefaultFile = 1,

        /// <summary>
        /// Do not use a default profile file.
        /// </summary>
        NoDefaultFile
    };

    /// <summary>
    /// Profile file create actions specify how to handle files that don't yet exist.
    /// </summary>
    public enum tvProfileFileCreateActions
    {
        /// <summary>
        /// Prompt the user to create the application's profile file (either
        /// the default profile file or a given alternative, created only if
        /// it doesn't already exist). Default settings will be automatically
        /// added to the profile file as they are encountered during the normal
        /// course of the application run.
        /// </summary>
        PromptToCreateFile = 1,

        /// <summary>
        /// Do not automatically create a profile file.
        /// </summary>
        NoFileCreate,

        /// <summary>
        /// Automatically create a profile file without user prompts
        /// (ie. "no questions asked").
        /// </summary>
        NoPromptCreateFile
    };

    /// <summary>
    /// Profile load actions specify how to handle items as they are loaded.
    /// </summary>
    public enum tvProfileLoadActions
    {
        /// <summary>
        /// Append all loaded items to the end of the profile.
        /// Duplicate keys are OK.
        /// </summary>
        Append = 1,

        /// <summary>
        /// Merge all loaded items into the profile. Matching items
        /// (by key) found in the profile will be replaced. Items
        /// to be loaded that do not match any current items will be appended
        /// to the end of the profile. Note: "*" as well as formal regular
        /// expressions can be used to match multiple keys. This means that
        /// long keys can be referenced on the command-line with "*"
        /// (eg. "-Auto*" matches "-AutoRun" and "-AutoPlay"). Keys with
        /// wildcards will not be appended.
        /// </summary>
        Merge,

        /// <summary>
        /// Clear the profile, then append all loaded items.
        /// </summary>
        Overwrite
    };


    /// <summary>
    /// <p>
    /// This class provides a simple yet flexible interface to
    /// a collection of persistent application level properties.
    /// </p>
    /// <p>
    /// The key feature of tvProfile is its seamless integration of
    /// file based properties with command-line arguments and switches.
    /// </p>
    /// <p>
    /// Application parameters (eg. <see langword='-Key1="value one" -Key2=abc
    /// -Key3=3'/>) can be intermixed with switches (eg. <see langword='-Switch'
    /// />) either in the application's profile file or on the command
    /// line or both. A switch is just shorthand for a boolean parameter. For
    /// example, <see langword='-Switch'/> is equivalent to <see langword=
    /// '-Switch=True'/>.
    /// </p>
    /// <p>
    /// tvProfile command-line switches / parameters typically override
    /// corresponding keys found in a profile file.
    /// </p>
    /// <p>
    /// Each application has a default plain text profile file named
    /// AppName.exe.txt, where AppName.exe is the executable filename of
    /// the application. The default profile file will always be found
    /// in the same folder as the application executable file. In fact,
    /// if it doesn't already exist in its default location, the application
    /// will automatically create the default profile file and automatically
    /// populate it with default values.
    /// </p>
    /// <p>
    /// As an alternative to the default delimited "command-line" profile
    /// file format, XML can be used instead by passing a boolean to the
    /// class constructor or by setting <see cref="bUseXmlFiles"/>.
    /// </p>
    /// <p>
    /// See <see langword='Remarks'/> for additional options.
    /// </p>
    /// <p>
    /// Author:     George Schiro (GeoCode@Schiro.name)
    /// </p>
    /// <p>
    /// Version:    2.28
    /// Copyright:  1996 - 2121
    /// </p>
    /// </summary>
    /// <remarks>
    /// The following are all of the "predefined" profile parameters (ie. those
    /// that exist for every profile):
    ///
    /// <list type="table">
    /// <listheader>
    ///     <term>Parameter</term>
    ///     <description>Description</description>
    /// </listheader>
    /// <item>
    ///     <term>-ini="path/file"</term>
    ///     <description>
    ///     A profile file location used to override the default profile file.
    ///     Alternatively, if the first argument passed on the command-line is
    ///     actually a file location (ie. a path/file specification) that refers
    ///     to an existing file, that file will be assumed to be a profile file
    ///     to override the default.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-File="path/file"</term>
    ///     <description>
    ///     The first command-line argument passed to the application, if a
    ///     profile file location has otherwise been provided. In other words,
    ///     if there is already an "-ini" key passed on the command-line
    ///     (after the first argument), then any file passed as the first
    ///     argument (a file that actually exists) will be added to the
    ///     profile as <see langword='-File="path/file"' />.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-NoCreate</term>
    ///     <description>
    ///     False by default. This and "-ini" (with its alias -ProfileFile) are
    ///     the only parameters that do not appear in a profile file. It is only
    ///     passed on the application command-line. When the switch <see langword='-NoCreate'/>
    ///     appears on the command-line, users are not prompted to create a profile
    ///     file. If a default profile does not exist at runtime, one will not
    ///     be created and default values will not be persisted. This option only
    ///     makes sense when the application needs to run with its original
    ///     "built-in" default values only.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-ProfileFile="path/file"</term>
    ///     <description>
    ///     A file used to override the default profile file (same as <see langword='-ini' />).
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-SaveProfile</term>
    ///     <description>
    ///     True by default (after the profile has been loaded from
    ///     a text file). Set this false to prevent automated changes
    ///     to the profile file.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-SaveSansCmdLine</term>
    ///     <description>
    ///     True by default. Set this false to allow merged command-lines
    ///     to be written to the profile file. When true, everything but
    ///     command-line keys will be saved.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-ShowProfile</term>
    ///     <description>
    ///     False by default. Display the contents of the profile in "command-line"
    ///     format during application startup. This is sometimes helpful for
    ///     debugging purposes.
    ///     </description>
    /// </item>
    /// <item>
    ///     <term>-XML_Profile</term>
    ///     <description>
    ///     False by default. Set this true to convert the profile file to XML
    ///     format and to maintain it that way. Set it false to convert it
    ///     back to line delimited "command-line" format.
    ///     </description>
    /// </item>
    /// </list>
    ///
    /// </remarks>
    public class tvProfile : ArrayList
    {
        #region "Constructors, Statics and Overridden or Augmented Members"

        /// <summary>
        /// Initializes a new instance of the tvProfile class.
        /// </summary>
        public tvProfile() : this(
                                      tvProfileDefaultFileActions.NoDefaultFile
                                    , tvProfileFileCreateActions.NoFileCreate
                                    )
        {
        }

        /// <summary>
        /// This is the main constructor.
        ///
        /// This constructor (or one of its shortcuts) is typically used at
        /// the top of an application during initialization.
        ///
        /// The profile is first initialized using the aeDefaultFileAction
        /// enum. Then any command-line arguments (typically passed from
        /// the environment) are merged into the profile. This way command
        /// line arguments override properties in the profile file.
        ///
        /// Initialization of the profile is done by first loading
        /// properties from an existing default profile file.
        ///
        /// The aeFileCreateAction enum enables a new profile file to be
        /// created (with or without prompting), if it doesn't already exist
        /// and then filled with default settings.
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        /// <param name="aeDefaultFileAction">
        /// This enum indicates how to handle automatic loading and saving
        /// of the default profile file.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the default profile file, if it doesn't already exist.
        /// </param>
        /// <param name="abUseXmlFiles">
        /// If true, XML file format will be used. If false, line delimited
        /// "command-line" format will be used (the default format).
        /// </param>
        public tvProfile(

                  string[]                      asCommandLineArray
                , tvProfileDefaultFileActions   aeDefaultFileAction
                , tvProfileFileCreateActions    aeFileCreateAction
                , bool                          abUseXmlFiles
                )
        {
            this.sInputCommandLineArray = asCommandLineArray;
            this.eDefaultFileAction = aeDefaultFileAction;
            this.eFileCreateAction = aeFileCreateAction;
            this.bUseXmlFiles = abUseXmlFiles;

            this.ReplaceDefaultProfileFromCommandLine(asCommandLineArray);
            if ( this.bExit )
            {
                return;
            }

            // "bDefaultFileReplaced = True" means that a replacement profile file has been passed on
            // the command-line. Consequently, no attempt to load the default profile file should be made.
            if ( !this.bDefaultFileReplaced  )
            {
                if ( tvProfileDefaultFileActions.NoDefaultFile != aeDefaultFileAction )
                {
                    this.Load(null, tvProfileLoadActions.Overwrite);
                    if ( this.bExit )
                    {
                        return;
                    }
                }

                this.LoadFromCommandLineArray(asCommandLineArray, tvProfileLoadActions.Merge);
            }

            bool    lbShowProfile = false;
                    if ( this.bAddStandardDefaults || this.ContainsKey("-ShowProfile") )
                    {
                        lbShowProfile = this.bValue("-ShowProfile", false);
                    }

            if ( lbShowProfile && null != tvProfile.oMsgBoxShow3 && /* DialogResult.Cancel */ 2 == (int)tvProfile.oMsgBoxShow3.Invoke(null
                        , new object[]{this.sCommandLine(), this.sLoadedPathFile, 1 /* MessageBoxButtons.OKCancel */ }) )
            {
                this.bExit = true;
            }
        }

        /// <summary>
        /// This is the main constructor used with abitrary XML documents.
        ///
        /// This constructor (or one of its shortcuts) is typically used at
        /// the top of an application during initialization.
        ///
        /// The profile is first initialized using the aeDefaultFileAction
        /// enum. Then any command-line arguments (typically passed from
        /// the environment) are merged into the profile. This way command
        /// line arguments override properties in the profile file.
        ///
        /// Initialization of the profile is done by first loading
        /// properties from an existing default profile file.
        ///
        /// The aeFileCreateAction enum enables a new profile file to be
        /// created (with or without prompting), if it doesn't already exist
        /// and then filled with default settings.
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        /// <param name="aeDefaultFileAction">
        /// This enum indicates how to handle automatic loading and saving
        /// of the default profile file.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the default profile file, if it doesn't already exist.
        /// </param>
        /// <param name="asXmlFile">
        /// This is the path/file that contains the XML to load.
        /// </param>
        /// <param name="asXmlXpath">
        /// This is the Xpath into asXmlFile that contains the profile.
        /// </param>
        /// <param name="asXmlKeyKey">
        /// This is the "Key" key used to find name attributes in asXmlXpath.
        /// </param>
        /// <param name="asXmlValueKey">
        /// This is the "Value" key used to find value attributes in asXmlXpath.
        /// </param>
        public tvProfile(

                  string[]                      asCommandLineArray
                , tvProfileDefaultFileActions   aeDefaultFileAction
                , tvProfileFileCreateActions    aeFileCreateAction
                , string                        asXmlFile
                , string                        asXmlXpath
                , string                        asXmlKeyKey
                , string                        asXmlValueKey
                )
        {
            this.sInputCommandLineArray = asCommandLineArray;
            this.eDefaultFileAction = aeDefaultFileAction;
            this.eFileCreateAction = aeFileCreateAction;
            this.sXmlXpath = asXmlXpath;
            this.sXmlKeyKey = asXmlKeyKey;
            this.sXmlValueKey = asXmlValueKey;
            this.bUseXmlFiles = true;

            this.ReplaceDefaultProfileFromCommandLine(asCommandLineArray);
            if ( this.bExit )
            {
                return;
            }

            // "bDefaultFileReplaced = True" means that a replacement profile file has been passed on
            // the command-line. Consequently, no attempt to load the default profile file should be made.
            if ( !this.bDefaultFileReplaced )
            {
                if ( tvProfileDefaultFileActions.NoDefaultFile != aeDefaultFileAction )
                {
                    this.Load(asXmlFile, tvProfileLoadActions.Overwrite);
                    if ( this.bExit )
                    {
                        return;
                    }
                }

                this.LoadFromCommandLineArray(asCommandLineArray, tvProfileLoadActions.Merge);
            }

            bool    lbShowProfile = false;
                    if ( this.bAddStandardDefaults || this.ContainsKey("-ShowProfile") )
                    {
                        lbShowProfile = this.bValue("-ShowProfile", false);
                    }

            if ( lbShowProfile && null != tvProfile.oMsgBoxShow3 && /* DialogResult.Cancel */ 2 == (int)tvProfile.oMsgBoxShow3.Invoke(null
                        , new object[]{this.sCommandLine(), this.sLoadedPathFile, 1 /* MessageBoxButtons.OKCancel */}) )
            {
                this.bExit = true;
            }
        }

        /// <summary>
        /// This constructor is typically used to load or create a new
        /// profile file separate from the default profile file for the
        /// application.
        /// </summary>
        /// <param name="asPathFile">
        /// This is the location of the profile file to be loaded.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the profile file to be loaded, if it doesn't already exist.
        /// </param>
        /// <param name="abUseXmlFiles">
        /// If true, XML file format will be used. If false, line delimited
        /// "command-line" format will be used (the default format).
        /// </param>
        public tvProfile(

                  string                        asPathFile
                , tvProfileFileCreateActions    aeFileCreateAction
                , bool                          abUseXmlFiles
                )
                : this()
        {
            this.eDefaultFileAction = tvProfileDefaultFileActions.NoDefaultFile;
            this.eFileCreateAction = aeFileCreateAction;
            this.bUseXmlFiles = abUseXmlFiles;

            this.Load(asPathFile, tvProfileLoadActions.Overwrite);
        }

        /// <summary>
        /// This constructor is the shortcut to the main constructor
        /// using the default profile file format.
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        /// <param name="aeDefaultFileAction">
        /// This enum indicates how to handle automatic loading and saving
        /// of the default profile file.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the default profile file, if it doesn't already exist.
        /// </param>
        public tvProfile(

                  string[]                      asCommandLineArray
                , tvProfileDefaultFileActions   aeDefaultFileAction
                , tvProfileFileCreateActions    aeFileCreateAction
                )
                : this(   asCommandLineArray
                        , aeDefaultFileAction
                        , aeFileCreateAction
                        , false
                        )
        {
        }

        /// <summary>
        /// This is the minimal profile constructor.
        ///
        /// This constructor is typically used to create an empty profile
        /// object to be populated manually (ie. with <see cref="Add(string, object)"/> calls).
        ///
        /// It is also used in lieu of the main constructor when command
        /// line overrides are not permitted.
        ///
        /// The default constructor is a shortcut to this one using the arguments
        /// "NoDefaultFile" and "NoFileCreate".
        /// </summary>
        /// <param name="aeDefaultFileAction">
        /// This enum indicates how to handle automatic loading and saving
        /// of the default profile file.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the default profile file, if it doesn't already exist.
        /// </param>
        public tvProfile(

                  tvProfileDefaultFileActions aeDefaultFileAction
                , tvProfileFileCreateActions aeFileCreateAction
                )
        {
            this.eDefaultFileAction = aeDefaultFileAction;
            this.eFileCreateAction = aeFileCreateAction;

            if ( tvProfileDefaultFileActions.AutoLoadSaveDefaultFile == aeDefaultFileAction )
                this.Load(this.sDefaultPathFile, tvProfileLoadActions.Overwrite);
        }

        /// <summary>
        /// This constructor initializes a profile object from a command-line
        /// string (eg. <see langword='-Key1="value one" -Switch1 -Key2=2'/>)
        /// rather than from a string array.
        ///
        /// This is handy for building profiles from simple strings passed
        /// from a database, from other profiles or from strings embedded
        /// in the body of an application.
        /// </summary>
        /// <param name="asCommandLine">
        /// This string (not a string array) should contain a command-line
        /// of <see langword='-key="value"'/> pairs.
        /// </param>
        public tvProfile(string asCommandLine) : this()
        {
            this.sInputCommandLine = asCommandLine;

            this.LoadFromCommandLine(asCommandLine, tvProfileLoadActions.Overwrite);
        }

        /// <summary>
        /// This constructor is typically used to load or create a new
        /// profile file separate from the default profile file for the
        /// application (using the default profile file format).
        /// </summary>
        /// <param name="asPathFile">
        /// This is the location of the profile file to be loaded.
        /// </param>
        /// <param name="aeFileCreateAction">
        /// This enum indicates how to handle the automatic creation of
        /// the profile file to be loaded, if it doesn't already exist.
        /// </param>
        public tvProfile(

                  string                        asPathFile
                , tvProfileFileCreateActions    aeFileCreateAction
                )
                : this(   asPathFile
                        , aeFileCreateAction
                        , false
                        )
        {
        }

        /// <summary>
        /// This constructor is typically used to load a profile file 
        /// separate from the default profile file for the application.
        /// The profile file must already exist. It will not be created.
        /// </summary>
        /// <param name="asPathFile">
        /// This is the location of the profile file to be loaded.
        /// </param>
        /// <param name="abUseXmlFiles">
        /// If true, XML file format will be used. If false, line delimited
        /// "command-line" format will be used (the default format).
        /// </param>
        public tvProfile(

                  string    asPathFile
                , bool      abUseXmlFiles
                )
                : this(   asPathFile
                        , tvProfileFileCreateActions.NoFileCreate
                        , abUseXmlFiles
                        )
        {
        }

        /// <summary>
        /// This constructor is the shortcut to the main XML constructor
        /// using the most typical options (ie. "AutoLoadSaveDefaultFile"
        /// and "PromptToCreateFile").
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        /// <param name="asXmlFile">
        /// This is the path/file that contains the XML to load.
        /// </param>
        /// <param name="asXmlXpath">
        /// This is the Xpath into asXmlFile that contains the profile.
        /// </param>
        /// <param name="asXmlKeyKey">
        /// This is the "Key" key used to find name attributes in asXmlXpath.
        /// </param>
        /// <param name="asXmlValueKey">
        /// This is the "Value" key used to find value attributes in asXmlXpath.
        /// </param>
        public tvProfile(

                  string[]  asCommandLineArray
                , string    asXmlFile
                , string    asXmlXpath
                , string    asXmlKeyKey
                , string    asXmlValueKey
                )
                : this(   asCommandLineArray
                        , tvProfileDefaultFileActions.AutoLoadSaveDefaultFile
                        , tvProfileFileCreateActions.PromptToCreateFile
                        , asXmlFile
                        , asXmlXpath
                        , asXmlKeyKey
                        , asXmlValueKey
                        )
        {
        }

        /// <summary>
        /// This constructor is the shortcut to the main constructor
        /// using the most typical options (ie. "AutoLoadSaveDefaultFile"
        /// and "PromptToCreateFile").
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        /// <param name="abUseXmlFiles">
        /// If true, XML file format will be used. If false, line delimited
        /// "command-line" format will be used (the default format).
        /// </param>
        public tvProfile(

                  string[]  asCommandLineArray
                , bool      abUseXmlFiles
                )
                : this(   asCommandLineArray
                        , tvProfileDefaultFileActions.AutoLoadSaveDefaultFile
                        , tvProfileFileCreateActions.PromptToCreateFile
                        , abUseXmlFiles
                        )
        {
        }

        /// <summary>
        /// This constructor is the shortcut to the main constructor
        /// using the most typical options (ie. "AutoLoadSaveDefaultFile",
        /// "PromptToCreateFile" and the default profile file format).
        /// </summary>
        /// <param name="asCommandLineArray">
        /// This string array is typically passed from the environment to the
        /// running application (eg. from Environment.GetCommandLineArgs() ).
        /// It is merged with the default profile file or any other profile
        /// found within the list of command-line arguments.
        /// </param>
        public tvProfile(

                string[] asCommandLineArray
                )
                : this(   asCommandLineArray
                        , tvProfileDefaultFileActions.AutoLoadSaveDefaultFile
                        , tvProfileFileCreateActions.PromptToCreateFile
                        , false
                        )
        {
        }


        /// <summary>
        /// This returns the global tvProfile object.
        /// 
        /// Note: use "tvProfile.oGlobal(aoProfile)" to set it.
        /// </summary>
        /// <returns>The global tvProfile object.</returns>
        public static tvProfile oGlobal()
        {
            return goGlobal;
        }

        /// <summary>
        /// This sets, then returns the global tvProfile object.
        /// 
        /// Note: use "tvProfile.oGlobal()" to get it without setting it.
        /// </summary>
        /// <returns>The global tvProfile object.</returns>
        public static tvProfile oGlobal(tvProfile aoProfile)
        {
            goGlobal = aoProfile;

            return goGlobal;
        }
        private static tvProfile goGlobal;


        #region "SortedList Member Emulation and Other Overrides"

        // The following methods don't necessarily augment or override ArrayList
        // members. They allow this class to emulate SortedList with added support
        // for duplicate keys. Inherit from SortedList and comment out these members.
        // Then this class will behave like SortedList.

        /// <summary>
        /// Adds the given "key/value" pair to the end of the profile.
        /// </summary>
        /// <param name="asKey">
        /// The key string stored with the value object.
        /// </param>
        /// <param name="aoValue">
        /// The value (as a generic object) to add.
        /// </param>
        public void Add(string asKey, object aoValue)
        {
            DictionaryEntry loEntry = new DictionaryEntry();
                            loEntry.Key = asKey;
                            loEntry.Value = aoValue;

            base.Add(loEntry);
        }

        /// <summary>
        /// Overrides the base <see langword='Add(object)'/> method in ArrayList.
        /// It throws an exception if the given object is not a DictionaryEntry.
        /// </summary>
        /// <param name="aoDictionaryEntry">
        /// The DictionaryEntry object to add to the collection.
        /// </param>
        /// <returns>
        /// The System.Collections.ArrayList index at which the object has been added.
        /// </returns>
        public override int Add(object aoDictionaryEntry)
        {
            if ( typeof(DictionaryEntry) != aoDictionaryEntry.GetType() )
            {
                throw new InvalidAddType();
            }
            else
            {
                return base.Add(aoDictionaryEntry);
            }
        }

        /// <summary>
        /// Returns true if the given object value is found in the profile.
        /// </summary>
        /// <param name="aoValue">
        /// The object to look for. The objects searched will be converted
        /// to strings if aoValue is a string. "*" or a regular expression
        /// may be included in aoValue.
        /// </param>
        /// <returns>
        /// True if found, false if not.
        /// </returns>
        public override bool Contains(object aoValue)
        {
            if ( "".GetType() == aoValue.GetType() )
            {
                string lsValue = (null == aoValue ? "" : aoValue.ToString());

                if ( mbUseLiteralsOnly )
                {
                    foreach ( DictionaryEntry loEntry in this )
                    {
                        if ( lsValue == (null == loEntry.Value ? "" : loEntry.Value.ToString()) )
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    foreach ( DictionaryEntry loEntry in this )
                    {
                        string lsExpression = this.sExpression(lsValue);

                        if ( null != lsExpression && Regex.IsMatch(null == loEntry.Value ? "" : loEntry.Value.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                        {
                            return true;
                        }
                    }
                }
            }
            else
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    if ( loEntry.Value == aoValue )
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Returns true if the given literal string value is found in the profile.
        /// </summary>
        /// <param name="asValue">
        /// The string value to look for. Regular expressions are ignored
        /// (ie. all characters in asValue are treated as literals).
        /// </param>
        /// <returns>
        /// True if found, false if not.
        /// </returns>
        public bool ContainsLiteral(string asValue)
        {
            foreach ( DictionaryEntry loEntry in this )
            {
                if ( asValue == (null == loEntry.Value ? "" : loEntry.Value.ToString()) )
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Returns true if the given key string is found in the profile.
        /// </summary>
        /// <param name="asKey">
        /// The key string to look for. "*" or a regular expression may be
        /// included.
        /// </param>
        /// <returns>
        /// True if found, false if not.
        /// </returns>
        public bool ContainsKey(string asKey)
        {
            foreach ( DictionaryEntry loEntry in this )
            {
                if ( mbUseLiteralsOnly )
                {
                    if ( asKey == (null == loEntry.Key ? "" : loEntry.Key.ToString()) )
                    {
                        return true;
                    }
                }
                else
                {
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(null == loEntry.Key ? "" : loEntry.Key.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// The zero-based index of the first entry in the profile with a key
        /// that matches asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string to look for. "*" or a regular expression may be
        /// included.
        /// </param>
        /// <returns>
        /// The zero-based index of the entry found. -1 is returned if no entry is found.
        /// </returns>
        public int IndexOfKey(string asKey)
        {
            if ( mbUseLiteralsOnly )
            {
                for ( int i = 0; i <= this.Count - 1; i++ )
                {
                    if ( asKey == (null == ((DictionaryEntry) base[i]).Key ? "" : ((DictionaryEntry) base[i]).Key.ToString()) )
                    {
                        return i;
                    }
                }
            }
            else
            {
                for ( int i = 0; i <= this.Count - 1; i++ )
                {
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(null == ((DictionaryEntry) base[i]).Key ? "" : ((DictionaryEntry) base[i]).Key.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        /// <summary>
        /// The zero-based index of the first entry in the profile with a value
        /// that matches asValue.
        /// </summary>
        /// <param name="asValue">
        /// The string value to look for. The objects searched will be converted
        /// to strings. "*" or a regular expression may be included in asValue.
        /// </param>
        /// <returns>
        /// The zero-based index of the entry found. -1 is returned if no entry is found.
        /// </returns>
        public int IndexOf(string asValue)
        {
            if ( mbUseLiteralsOnly )
            {
                for ( int i = 0; i <= this.Count - 1; i++ )
                {
                    if ( asValue == (null == ((DictionaryEntry) base[i]).Value ? "" : ((DictionaryEntry) base[i]).Value.ToString()) )
                    {
                        return i;
                    }
                }
            }
            else
            {
                for ( int i = 0; i <= this.Count - 1; i++ )
                {
                    string lsExpression = this.sExpression(asValue);

                    if ( null != lsExpression && Regex.IsMatch((null == ((DictionaryEntry) base[i]).Value ? "" : ((DictionaryEntry) base[i]).Value.ToString()), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        /// <summary>
        /// This is the string indexer for the class.
        /// </summary>
        /// <param name="asKey">
        /// The key string to look for. "*" or a regular expression may be
        /// included.
        /// </param>
        /// <returns>
        /// Gets or sets the object value of the first entry found that matches asKey.
        /// </returns>
        public object this[string asKey]
        {
            get
            {
                int liIndex = this.IndexOfKey(asKey);

                if ( -1 == liIndex )
                {
                    return null;
                }
                else
                {
                    return this[liIndex];
                }
            }
            set
            {
                int liIndex = this.IndexOfKey(asKey);

                if ( -1 == liIndex )
                {
                    this.Add(asKey, value);
                }
                else
                {
                    this.SetByIndex(liIndex, value);
                }
            }
        }

        /// <summary>
        /// This is the integer indexer for the class.
        /// </summary>
        /// <param name="aiIndex">
        /// The zero-based index to look for.
        /// </param>
        /// <returns>
        /// Gets or sets the object value of the entry found at the given zero-based index position.
        /// </returns>
        public override object this[int aiIndex]
        {
            get
            {
                return ((DictionaryEntry) base[aiIndex]).Value;
            }
            set
            {
                this.SetByIndex(aiIndex, value);
            }
        }

        /// <summary>
        /// Removes zero, one or many entries in the profile with keys
        /// that match the given aoKey. aoKey will be cast as a string.
        /// </summary>
        /// <param name="aoKey">
        /// The key string to look for. "*" or a regular expression may be
        /// included.
        /// </param>
        public override void Remove(object aoKey)
        {
            this.Remove(aoKey.ToString());
        }

        /// <summary>
        /// Removes zero, one or many entries in the profile with keys
        /// that match the given asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string to look for. "*" or a regular expression may be
        /// included.
        /// </param>
        public void Remove(string asKey)
        {
            int liIndex;

            do
            {
                liIndex = this.IndexOfKey(asKey);

                if ( -1 != liIndex )
                {
                    this.RemoveAt(liIndex);
                }

            } while ( -1 != liIndex );
        }

        /// <summary>
        /// Sets the object at the given zero-based index position within the
        /// profile to the given object value. This method is called by the
        /// integer indexer for the class.
        /// </summary>
        /// <param name="aiIndex">
        /// The zero-based index to look for.
        /// </param>
        /// <param name="aoValue">
        /// The object value that is written to the entry at the given zero-based index position.
        /// </param>
        public void SetByIndex(int aiIndex, object aoValue)
        {
            string lsKey = (null == ((DictionaryEntry) base[aiIndex]).Key ? "" : ((DictionaryEntry) base[aiIndex]).Key.ToString());

            base[aiIndex] = new DictionaryEntry(lsKey, aoValue);
        }

        /// <summary>
        /// Sorts the the profile by its key strings
        /// (alias of SortByKey).
        /// </summary>
        public override void Sort()
        {
            this.SortByKey();
        }

        /// <summary>
        /// Sorts the the profile by its key strings.
        /// </summary>
        public void SortByKey()
        {
            this.Sort(new tvProfileKeyComparer());
        }

        /// <summary>
        /// Sorts the the profile by its values (as strings).
        /// </summary>
        public void SortByValueAsString()
        {
            this.Sort(new tvProfileStringValueComparer());
        }

        /// <summary>
        /// Overrides base ToString() method.
        /// </summary>
        /// <returns>tvProfile contents in command-line block format.</returns>
        public override string ToString()
        {
            return this.sCommandBlock();
        }
        #endregion

        #endregion

        /// <summary>
        /// Returns true if the standard "built-in" profile defaults
        /// will be automatically added to the profile. This property
        /// will generally be true within the main constructors,
        /// unless no default profile file is used.
        /// </summary>
        public  bool  bAddStandardDefaults
        {
            get
            {
                if ( !mbAddStandardDefaults_init )
                {
                    if ( tvProfileDefaultFileActions.AutoLoadSaveDefaultFile == this.eDefaultFileAction )
                        mbAddStandardDefaults = true;

                    mbAddStandardDefaults_init = true;
                }

                return mbAddStandardDefaults;
            }
            set
            {
                mbAddStandardDefaults = value;
                mbAddStandardDefaults_init = true;
            }
        }
        private bool mbAddStandardDefaults = false;
        private bool mbAddStandardDefaults_init = false;

        /// <summary>
        /// Returns true if the profile was instanced from a file loaded using
        /// the predefined parameter: <see langword='-ini="path/file"'/>
        /// (see <see cref="tvProfile"/> remarks). In other words, this property
        /// returns true if the application's default profile file was replaced.
        /// </summary>
        public  bool  bDefaultFileReplaced
        {
            get
            {
                return mbDefaultFileReplaced;
            }
            set
            {
                mbDefaultFileReplaced = value;
            }
        }
        private bool mbDefaultFileReplaced = false;

        /// <summary>
        /// Returns true if the profile file is maintained in a
        /// locked state while the profile object exists. This
        /// prevents overstepping access by external processes.
        /// 
        /// The default value of this property is true.
        /// </summary>
        public  bool  bEnableFileLock
        {
            get
            {
                return mbEnableFileLock;
            }
            set
            {
                mbEnableFileLock = value;

                if ( mbEnableFileLock )
                    this.bLockProfileFile(this.sActualPathFile);
                else
                    this.UnlockProfileFile();
            }
        }
        private bool mbEnableFileLock = true;

        /// <summary>
        /// Returns true if the user selected "Cancel" in response to a profile message.
        /// </summary>
        public  bool  bExit
        {
            get
            {
                return mbExit;
            }
            set
            {
                mbExit = value;

                if ( mbExit )
                {
                    this.bSaveEnabled = false;
                }
            }
        }
        private bool mbExit = false;

        /// <summary>
        /// Returns true if the profile's file was just created. It returns
        /// false if the profile file existed previously (at runtime).
        /// </summary>
        public  bool  bFileJustCreated
        {
            get
            {
                return mbFileJustCreated;
            }
            set
            {
                mbFileJustCreated = value;
            }
        }
        private bool mbFileJustCreated = false;

        /// <summary>
        /// Returns true if the EXE is in a folder (or subfolder) with the same name.
        /// </summary>
        public bool bInOwnFolder
        {
            get
            {
                string lsPathOnly = Path.GetDirectoryName(this.sExePathFile);
                string lsFilnameOnly = Path.GetFileNameWithoutExtension(this.sExePathFile);

                return -1 != lsPathOnly.ToLower().IndexOf(lsFilnameOnly.ToLower());
            }
        }

        /// <summary>
        /// Returns true if the EXE is already in a typical installation folder
        /// (eg. "Program Files").
        /// </summary>
        public bool bInstalledAlready
        {
            get
            {
                string[] lsPathFragArray = {
                          Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
                        , Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
                        , Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))
                        , Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
                        , Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
                        , Path.GetFileName(Environment.GetFolderPath(Environment.SpecialFolder.Programs))
                        };

                foreach (string lsPathFrag in lsPathFragArray)
                {
                    if ( -1 != this.sExePathFile.IndexOf(lsPathFrag) )
                        return true;
                }

                return false;
            }
        }

        /// <summary>
        /// <p>
        /// Returns true if the <see cref="Save()"/> method is enabled. <see cref="Save()"/>
        /// is enabled after a text file is successfully loaded or a text file
        /// is saved with <c>Save(asPathFile)</c>.
        /// </p>
        /// <p>
        /// The predefined <see langword='-SaveProfile=false'/> switch can
        /// be used to disable the <see cref="Save()"/> method manually
        /// (see <see cref="tvProfile"/> remarks).
        /// </p>
        /// </summary>
        public  bool  bSaveEnabled
        {
            get
            {
                if ( this.bAddStandardDefaults || this.ContainsKey(mcsSaveKey) )
                {
                    if ( mbSaveEnabled )
                    {
                        if ( mbInputCommandLineArraySaveEnabledSet )
                            mbSaveEnabled = mbInputCommandLineArraySaveEnabled;
                        else
                            mbSaveEnabled = this.bValue(mcsSaveKey, mbSaveEnabled);
                    }
                }

                return mbSaveEnabled;
            }
            set
            {
                mbSaveEnabled = value;
            }
        }
        private bool mbSaveEnabled = false;

        /// <summary>
        /// Returns true if all but command-line merged keys will be saved to a profile
        /// file. The predefined <see langword='-SaveSansCmdLine=false'/> switch can be
        /// used to disable this behavior so even merged profiles are saved (see 
        /// <see cref="tvProfile"/> remarks).
        /// </summary>
        public  bool  bSaveSansCmdLine
        {
            get
            {
                if ( this.bAddStandardDefaults || this.ContainsKey("-SaveSansCmdLine") )
                    mbSaveSansCmdLine = this.bValue("-SaveSansCmdLine", mbSaveSansCmdLine);

                if ( mbSaveSansCmdLine && null == moInputCommandLineProfile )
                {
                    moInputCommandLineProfile = new tvProfile();
                    moInputCommandLineProfile.LoadFromCommandLineArray(this.sInputCommandLineArray, tvProfileLoadActions.Append);

                    moMatchCommandLineProfile = new tvProfile();

                    foreach ( DictionaryEntry loEntry in moInputCommandLineProfile )
                    {
                        // This must be a "merge" here in case there is more than one of the same key on the command-line.
                        moMatchCommandLineProfile.LoadFromCommandLineArray(
                                this.oOneKeyProfile(loEntry.Key.ToString()).sCommandLineArray(), tvProfileLoadActions.Merge);
                    }
                }

                return mbSaveSansCmdLine;
            }
            set
            {
                mbSaveSansCmdLine = value;
            }
        }
        private bool mbSaveSansCmdLine = true;

        /// <summary>
        /// Returns true if profile files will be read and written in XML format
        /// rather than the default line delimited "command-line" format.
        /// 
        /// The default value of this property is false.
        /// </summary>
        public  bool  bUseXmlFiles
        {
            get
            {
                if ( this.bAddStandardDefaults || this.ContainsKey("-XML_Profile") )
                {
                    mbUseXmlFiles = this.bValue("-XML_Profile", mbUseXmlFiles);
                }

                return mbUseXmlFiles;
            }
            set
            {
                if ( value != mbUseXmlFiles )
                {
                    if ( this.bAddStandardDefaults || this.ContainsKey("-XML_Profile") )
                    {
                        this["-XML_Profile"] = value;
                    }

                    this.sLoadedPathFile = this.sReformatProfileFile(this.sLoadedPathFile);
                }

                mbUseXmlFiles = value;
            }
        }
        private bool mbUseXmlFiles = false;

        /// <summary>
        /// Returns true if all input strings are assumed to be literal while
        /// searching key strings and value strings (ie. no regular expressions).
        /// 
        /// The default value of this property is false.
        /// </summary>
        public bool bUseLiteralsOnly
        {
            get
            {
                return mbUseLiteralsOnly;
            }
            set
            {
                mbUseLiteralsOnly = value;
            }
        }
        private bool mbUseLiteralsOnly = false;

        /// <summary>
        /// The path/file location most recently used either to load the profile
        /// from a text file or to save the profile to a text file.
        /// </summary>
        public  string  sActualPathFile
        {
            get
            {
                return msActualPathFile;
            }
            set
            {
                msActualPathFile = value;

                if ( !String.IsNullOrEmpty(msActualPathFile) && String.IsNullOrEmpty(Path.GetPathRoot(msActualPathFile)) )
                {
                    string  lsCwdBasedPathFile = Path.Combine(Environment.CurrentDirectory, msActualPathFile);
                    string  lsExeBasedPathFile = this.sRelativeToExePathFile(msActualPathFile);
                            if ( File.Exists(lsCwdBasedPathFile) )
                                msActualPathFile = lsCwdBasedPathFile;
                            else
                            if ( File.Exists(lsExeBasedPathFile) )
                                msActualPathFile = lsExeBasedPathFile;
                            else
                                msActualPathFile = lsCwdBasedPathFile;
                }

                this.bSaveEnabled = true;
            }
        }
        private string msActualPathFile;

        /// <summary>
        /// The original default file action passed to the constructor.
        /// See <see cref="tvProfileDefaultFileActions"/>.
        /// </summary>
        public  tvProfileDefaultFileActions  eDefaultFileAction
        {
            get
            {
                return meDefaultFileAction;
            }
            set
            {
                meDefaultFileAction = value;

                mbAddStandardDefaults = false;
                mbAddStandardDefaults_init = false;
            }
        }
        private tvProfileDefaultFileActions meDefaultFileAction = tvProfileDefaultFileActions.NoDefaultFile;

        /// <summary>
        /// The original file create action passed to the constructor.
        /// See <see cref="tvProfileFileCreateActions"/>.
        /// </summary>
        public  tvProfileFileCreateActions  eFileCreateAction
        {
            get
            {
                if ( this.ContainsKey("-NoCreate") )
                {
                    // ".Contains" is used here so that this switch
                    // is not automatically added to the profile. It
                    // can only be added to the profile manually.
                    return tvProfileFileCreateActions.NoFileCreate;
                }
                else
                {
                    return meFileCreateAction;
                }
            }
            set
            {
                meFileCreateAction = value;
            }
        }
        private tvProfileFileCreateActions meFileCreateAction = tvProfileFileCreateActions.NoFileCreate;

        /// <summary>
        /// The path/file location of a text file just created
        /// as an unlocked backup of the current profile file.
        /// </summary>
        public  string  sBackupPathFile
        {
            get
            {
                string  lsPath = Path.GetDirectoryName(this.sLoadedPathFile);
                string  lsFilename = Path.GetFileName(this.sLoadedPathFile);
                string  lsExt = Path.GetExtension(this.sLoadedPathFile);
                string  lsBackupPathFile = Path.Combine(lsPath, lsFilename) + ".backup" + lsExt;

                bool lbSaveEnabled = this.bSaveEnabled;
                this.Save(lsBackupPathFile);
                this.sActualPathFile = this.sLoadedPathFile;
                this.bSaveEnabled = lbSaveEnabled;

                return lsBackupPathFile;
            }
        }

        /// <summary>
        /// The default file extension of the profile's text file. If
        /// <see cref="bUseXmlFiles"/> is true, this method returns ".config",
        /// otherwise it returns ".txt".
        /// </summary>
        public  string    sDefaultFileExt
        {
            get
            {
                if ( String.IsNullOrEmpty(msDefaultFileExt) )
                {
                    if ( !this.bUseXmlFiles )
                    {
                        return msDefaultFileExtArray[0];
                    }
                    else
                    {
                        return msDefaultFileExtArray[1];
                    }
                }
                else
                {
                    return msDefaultFileExt;
                }
            }
            set
            {
                msDefaultFileExt = value;
            }
        }
        private string   msDefaultFileExt;
        private string[] msDefaultFileExtArray = {mcsLoadSaveDefaultExtension, ".config"};

        /// <summary>
        /// The default path/file location of the profile's text file. This
        /// property uses <see cref="sExePathFile"/>.
        /// </summary>
        public string sDefaultPathFile
        {
            get
            {
                string  lsPathFile = this.sExePathFile + this.sDefaultFileExt;
                        if ( !File.Exists(lsPathFile) )
                            lsPathFile = Path.GetFileName(this.sExePathFile) + this.sDefaultFileExt;

                return lsPathFile;
            }
        }


        /// <summary>
        /// The path/file location of the executing application or assembly
        /// that uses the profile (including the name of the executable).
        /// This property is used by <see cref="sDefaultPathFile"/>. Setting
        /// this property allows a virtual assembly location to be used as an
        /// alternative to the actual location.
        /// </summary>
        public  string  sExePathFile
        {
            get
            {
                if ( String.IsNullOrEmpty(msExePathFile) )
                {
                    try
                    {
                        msExePathFile = Assembly.GetEntryAssembly().Location;
                    }
                    catch
                    {
                        msExePathFile = Assembly.GetExecutingAssembly().Location;
                    }
                }

                return msExePathFile;
            }
            set
            {
                msExePathFile = value;
            }
        }
        private string msExePathFile;

        /// <summary>
        /// The original "command-line string" input passed to the constructor.
        /// </summary>
        public  string  sInputCommandLine
        {
            get
            {
                return msInputCommandLine;
            }
            set
            {
                msInputCommandLine = value;
            }
        }
        private string msInputCommandLine;

        /// <summary>
        /// The original "command-line string array" input passed to the constructor.
        /// </summary>
        public  string[]  sInputCommandLineArray
        {
            get
            {
                return msInputCommandLineArray;
            }
            set
            {
                msInputCommandLineArray = value;

                string lsSaveKey = mcsSaveKey.ToLower();

                if ( null != msInputCommandLineArray )
                    foreach(string lsItem in msInputCommandLineArray)
                    {
                        string[] lsKeyValueArray = lsItem.ToLower().Split(mccAsnMark);

                        if ( lsKeyValueArray[0].Trim() == lsSaveKey )
                        {
                            mbInputCommandLineArraySaveEnabledSet = true;
                            mbInputCommandLineArraySaveEnabled = 1 == lsKeyValueArray.Length || "true" == lsKeyValueArray[1].Trim();
                            this[mcsSaveKey] = mbInputCommandLineArraySaveEnabled;
                            break;
                        }
                    }
            }
        }
        private bool     mbInputCommandLineArraySaveEnabledSet;
        private bool     mbInputCommandLineArraySaveEnabled;
        private string[] msInputCommandLineArray;

        /// <summary>
        /// The path/file location most recently used to load the profile from
        /// a text file.
        /// </summary>
        public  string  sLoadedPathFile
        {
            get
            {
                return msLoadedPathFile;
            }
            set
            {
                msLoadedPathFile = value;

                if ( !String.IsNullOrEmpty(msLoadedPathFile) && String.IsNullOrEmpty(Path.GetPathRoot(msLoadedPathFile)) )
                {
                    string  lsCwdBasedPathFile = Path.Combine(Environment.CurrentDirectory, msLoadedPathFile);
                    string  lsExeBasedPathFile = this.sRelativeToExePathFile(msLoadedPathFile);
                            if ( File.Exists(lsCwdBasedPathFile) )
                                msLoadedPathFile = lsCwdBasedPathFile;
                            else
                            if ( File.Exists(lsExeBasedPathFile) )
                                msLoadedPathFile = lsExeBasedPathFile;
                            else
                                msLoadedPathFile = lsCwdBasedPathFile;
                }

                this.sActualPathFile = value;
            }
        }
        private string msLoadedPathFile;

        /// <summary>
        /// The "new line" character passed in with the source data.
        /// If none is found, "Environment.NewLine" is used by default.
        /// </summary>
        public  string  sNewLine
        {
            get
            {
                return msNewLine;
            }
            set
            {
                msNewLine = value;
            }
        }
        private string msNewLine = Environment.NewLine;

        /// <summary>
        /// The key used to find the "key" attribute in XML "key/value" pairs
        /// (ie. "key").
        /// </summary>
        public  string  sXmlKeyKey
        {
            get
            {
                return msXmlKeyKey;
            }
            set
            {
                msXmlKeyKey = value;
            }
        }
        private string msXmlKeyKey = "key";

        /// <summary>
        /// The key used to find the "value" attribute in XML "key/value" pairs
        ///  (ie. "value").
        /// </summary>
        public  string  sXmlValueKey
        {
            get
            {
                return msXmlValueKey;
            }
            set
            {
                msXmlValueKey = value;
            }
        }
        private string msXmlValueKey = "value";

        /// <summary>
        /// The Xpath expression used to find the profile section in a given
        /// XML document. This expression must have a format similar to the
        /// default of "configuration/appSettings/add" (ie. no wildcards).
        /// </summary>
        public  string  sXmlXpath
        {
            get
            {
                return msXmlXpath;
            }
            set
            {
                msXmlXpath = value;
            }
        }
        private string msXmlXpath = "configuration/appSettings/add";

        /// <summary>
        /// The value of the item found in the profile for the given
        /// asKey, returned as a generic object. If asKey doesn't exist in the
        /// profile, it will be added with the given aoDefault object value.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value object in the profile.
        /// </param>
        /// <param name="aoDefault">
        /// The default value object added to the profile with asKey,
        /// if asKey can't be found.
        /// </param>
        /// <returns>
        /// Returns the object found for asKey or aoDefault (if asKey
        /// is not found).
        /// </returns>
        public object GetAdd(string asKey, object aoDefault)
        {
            object loValue;

            if ( this.ContainsKey(asKey) )
            {
                loValue = this[asKey];
            }
            else
            {
                loValue = aoDefault;
                this.Add(asKey, loValue);

                if ( tvProfileDefaultFileActions.NoDefaultFile != this.eDefaultFileAction )
                {
                    this.Save();
                }
            }

            return loValue;
        }

        #region "Various GetAdd Return Types"

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey, cast as a boolean.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="abDefault">
        /// The boolean value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The boolean value found or abDefault (see <see cref="GetAdd"/>).
        /// </returns>
        public bool bValue(string asKey, bool abDefault)
        {
            return Convert.ToBoolean(this.GetAdd(asKey, abDefault));
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey, cast as a double.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="adDefault">
        /// The double value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The double value found or adDefault (see <see cref="GetAdd"/>).
        /// </returns>
        public double dValue(string asKey, double adDefault)
        {
            return Convert.ToDouble(this.GetAdd(asKey, adDefault));
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey, cast as a DateTime.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="adtDefault">
        /// The DateTime value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The DateTime value found or adtDefault (see <see cref="GetAdd"/>).
        /// </returns>
        public DateTime dtValue(string asKey, DateTime adtDefault)
        {
            return Convert.ToDateTime(this.GetAdd(asKey, adtDefault));
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey, cast as an integer.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="aiDefault">
        /// The integer value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The integer value found or aiDefault (see <see cref="GetAdd"/>).
        /// </returns>
        public int iValue(string asKey, int aiDefault)
        {
            return Convert.ToInt32(this.GetAdd(asKey, aiDefault));
        }

        /// <summary>
        /// The profile entry object at the given zero-based index position.
        /// </summary>
        /// <param name="aiIndex">
        /// The zero based index into the list of profile entry objects.
        /// </param>
        /// <returns>
        /// The DictionaryEntry object found at aiIndex.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <p>aiIndex is less than zero.</p>
        /// <p></p>
        /// -or-
        /// <p></p>
        /// aiIndex is equal to or greater than <see cref="ArrayList.Count"/>.
        /// </exception>
        public DictionaryEntry oEntry(int aiIndex)
        {
            return ((DictionaryEntry) base[aiIndex]);
        }

        /// <summary>
        /// Set the profile entry object at the given zero-based index position.
        /// </summary>
        /// <param name="aiIndex">
        /// The zero based index into the list of profile entry objects.
        /// </param>
        /// <param name="Value">
        /// The DictionaryEntry object to be stored at aiIndex.
        /// </param>
        public void SetEntry(int aiIndex, DictionaryEntry Value)
        {
            base[aiIndex] = Value;
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="aoDefaultProfile">
        /// The tvProfile object value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The tvProfile object value found or aoDefaultProfile (see <see cref="GetAdd"/>).
        /// </returns>
        public tvProfile oProfile(string asKey, tvProfile aoDefaultProfile)
        {
            object  loProfile = this.GetAdd(asKey, aoDefaultProfile);
            object  loProfileCast = loProfile as tvProfile;
                    if ( null == loProfileCast )
                        loProfileCast = new tvProfile(null == loProfile ? "" : loProfile.ToString());

            return (tvProfile)loProfileCast;
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="asDefaultProfile">
        /// The tvProfile object value (converted from a command-line string)
        /// returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The tvProfile object value found or asDefaultProfile 
        /// (converted to a tvProfile object, see <see cref="GetAdd"/>).
        /// </returns>
        public tvProfile oProfile(string asKey, string asDefaultProfile)
        {
            object  loProfile = this.GetAdd(asKey, asDefaultProfile);
            object  loProfileCast = loProfile as tvProfile;
                    if ( null == loProfileCast )
                        loProfileCast = new tvProfile(null == loProfile ? "" : loProfile.ToString());

            return (tvProfile)loProfileCast;
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <returns>
        /// The tvProfile object value found or an empty tvProfile
        /// (nothing is added to the parent profile if nothing is found).
        /// </returns>
        public tvProfile oProfile(string asKey)
        {
            object  loProfile = this[asKey];
            object  loProfileCast = loProfile as tvProfile;
                    if ( null == loProfileCast )
                        loProfileCast = new tvProfile(null == loProfile ? "" : loProfile.ToString());

            return (tvProfile)loProfileCast;
        }

        /// <summary>
        /// The value found for aiIndex.
        /// </summary>
        /// <param name="aiIndex">
        /// The integer index used to find the corresponding value in the profile.
        /// </param>
        /// <returns>
        /// The tvProfile object value found or an empty tvProfile
        /// (nothing is added to the parent profile if nothing is found).
        /// </returns>
        public tvProfile oProfile(int aiIndex)
        {
            object  loProfile = this[aiIndex];
            object  loProfileCast = loProfile as tvProfile;
                    if ( null == loProfileCast )
                        loProfileCast = new tvProfile(null == loProfile ? "" : loProfile.ToString());

            return (tvProfile)loProfileCast;
        }

        /// <summary>
        /// The profile returned after decompressing btArrayZipped.
        /// </summary>
        /// <param name="abtArrayZipped">
        /// A byte array that represents a ZIP-compressed command-line formatted profile string.
        /// </param>
        /// <returns>
        /// The tvProfile object resulting from decompressing the given byte array.
        /// </returns>
        public static tvProfile oProfile(byte[] abtArrayZipped)
        {
            using (MemoryStream loMemoryStream = new MemoryStream())
            {
                // Define the final length of the returned profile string as whatever is
                // specified in the first mciIntSizeInBytes bytes of the compressed data.
                byte[] lbtArrayProfileAsString = new byte[BitConverter.ToInt32(abtArrayZipped, 0)];

                // Write the given byte array to loMemoryStream (skipping the first mciIntSizeInBytes bytes).
                loMemoryStream.Write(abtArrayZipped, mciIntSizeInBytes, abtArrayZipped.Length - mciIntSizeInBytes);
                loMemoryStream.Position = 0;

                using (GZipStream loGZipStream = new GZipStream(loMemoryStream, CompressionMode.Decompress))
                {
                    // Using loGZipStream, decompress loMemoryStream to lbtArrayProfileAsString.
                    loGZipStream.Read(lbtArrayProfileAsString, 0, lbtArrayProfileAsString.Length);
                }

                // Return the decompressed profile string as a profile object.
                return new tvProfile(Encoding.UTF8.GetString(lbtArrayProfileAsString));
            }
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="aoDefault">
        /// The object value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The object value found or aoDefault (see <see cref="GetAdd"/>).
        /// </returns>
        public object oValue(string asKey, object aoDefault)
        {
            return this.GetAdd(asKey, aoDefault);
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey,
        /// cast as a trimmed string.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="asDefault">
        /// The string value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The string value found or asDefault (see <see cref="GetAdd"/>). Any
        /// environment variables found in the return string are expanded and the
        /// return string is trimmed of leading and trailing spaces.
        /// </returns>
        public string sValue(string asKey, string asDefault)
        {
            return this.sValueNoTrim(asKey, asDefault).Trim();
        }

        /// <summary>
        /// The "<see cref="GetAdd"/> object" value found for asKey, cast as a string (not trimmed).
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find the corresponding value in the profile.
        /// </param>
        /// <param name="asDefault">
        /// The string value returned if asKey is not found.
        /// </param>
        /// <returns>
        /// The string value found or asDefault (see <see cref="GetAdd"/>). Any
        /// environment variables found in the return string are expanded and the
        /// return string is NOT trimmed of leading or trailing spaces.
        /// </returns>
        public string sValueNoTrim(string asKey, string asDefault)
        {
            object loValue = this.GetAdd(asKey, asDefault);
            string lsValue = (null == loValue ? "" : loValue.ToString());

            // The return limit of ExpandEnvironmentVariables is 32K. The difference allows for variable expansion.
            if ( 32000 > lsValue.Length )
            {
                return Environment.ExpandEnvironmentVariables(lsValue);
            }
            else
            {
                return lsValue;
            }
        }

        #endregion

        /// <summary>
        /// Compressed version (in ZIP format) of this profile (represented as a command-line string).
        /// 
        /// Decompress using tvProfile.oProfile(abtArrayZipped).
        /// </summary>
        /// <returns>Byte array of this profile ZIP-compressed from its command-line string.</returns>
        public byte[] btArrayZipped()
        {
            byte[]          lbtArrayZipped = null;
            MemoryStream    loMemoryStream = new MemoryStream();
            byte[]          lbtArrayProfileAsString = Encoding.UTF8.GetBytes(this.sCommandLine());
                            using (GZipStream loGZipStream = new GZipStream(loMemoryStream, CompressionMode.Compress, true))
                            {
                                // Using loGZipStream, compress lbtArrayProfileAsString to loMemoryStream.
                                loGZipStream.Write(lbtArrayProfileAsString, 0, lbtArrayProfileAsString.Length);
                            }
            byte[]          lbtArrayProfileZipped = new byte[loMemoryStream.Length];
                            // Now write loMemoryStream to lbtArrayProfileZipped (an intermediate byte array).
                            loMemoryStream.Position = 0;
                            loMemoryStream.Read(lbtArrayProfileZipped, 0, lbtArrayProfileZipped.Length);

            // Define the final array with the original profile string size as the first mciIntSizeInBytes bytes.
            lbtArrayZipped = new byte[mciIntSizeInBytes + lbtArrayProfileZipped.Length];

            // Copy the original profile string length to the first mciIntSizeInBytes bytes.
            Buffer.BlockCopy(BitConverter.GetBytes(lbtArrayProfileAsString.Length), 0, lbtArrayZipped, 0, mciIntSizeInBytes);

            // Copy the intermediate byte array to the new byte array (after the first first mciIntSizeInBytes bytes).
            Buffer.BlockCopy(lbtArrayProfileZipped, 0, lbtArrayZipped, mciIntSizeInBytes, lbtArrayProfileZipped.Length);

            return lbtArrayZipped;
        }

        /// <summary>
        /// The count of profile entries with a key that matches asKey.
        /// This number will be greater than one if asKey appears multiple
        /// times in the profile (duplicate keys are OK). Likewise, this
        /// number might be greater than one if asKey contains "*" or a
        /// regular expression.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find items in the profile. "*" or a regular
        /// expression may be included.
        /// </param>
        /// <returns>
        /// Integer count of the items found.
        /// </returns>
        public int iKeyCount(string asKey)
        {
            int liCount = 0;

            if ( mbUseLiteralsOnly )
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    if ( asKey == (null == loEntry.Key ? "" : loEntry.Key.ToString()) )
                    {
                        liCount++;
                    }
                }
            }
            else
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(null == loEntry.Key ? "" : loEntry.Key.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        liCount++;
                    }
                }
            }

            return liCount;
        }

        /// <summary>
        /// Returns the entire contents of the profile as a "command block" string
        /// (eg. <see langword='-Key1=one -Key2=2 -Key3 -Key4="we have four"'/>), 
        /// where the items are stacked vertically rather than listed horozontally.
        /// This feature is handy for minimally serializing profiles, passing them
        /// around easily and storing them elsewhere.
        /// </summary>
        /// <returns>
        /// A "command block" string (not a string array).
        /// </returns>
        public string sCommandBlock()
        {
            ++miCommandBlockRecursionLevel;

            string          lcsIndent = "".PadRight(4 * miCommandBlockRecursionLevel);
            StringBuilder   lsbCommandBlock = new StringBuilder(this.sNewLine);

            foreach ( DictionaryEntry loEntry in this )
            {
                string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());
                string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                if ( lsValue.Contains(this.sNewLine) )
                {
                    lsbCommandBlock.Append(lcsIndent + lsKey + mcsAsnMark + mcsBlockBegMark + this.sNewLine);
                    lsbCommandBlock.Append(lsValue);
                    lsbCommandBlock.Append(lcsIndent + lsKey + mcsAsnMark + mcsBlockEndMark + this.sNewLine);
                }
                else
                {
                    if ( lsValue.Contains(mcsSpcMark) || lsValue.Contains(mcsArgMark) )
                    {
                        lsbCommandBlock.Append(lcsIndent + lsKey + mcsAsnMark + mcsQteMark1 + lsValue + mcsQteMark1 + this.sNewLine);
                    }
                    else
                    {
                        lsbCommandBlock.Append(lcsIndent + lsKey + mcsAsnMark + lsValue + this.sNewLine);
                    }
                }
            }

            lsbCommandBlock.Append(this.sNewLine);

            --miCommandBlockRecursionLevel;

            return lsbCommandBlock.ToString();
        }
        private static int miCommandBlockRecursionLevel = 0;

        /// <summary>
        /// Returns the entire contents of the profile as a "command-line" string
        /// (eg. <see langword='-Key1=one -Key2=2 -Key3 -Key4="we have four"'/>).
        /// This feature is handy for minimally serializing profiles, passing them
        /// around easily and storing them elsewhere.
        /// </summary>
        /// <returns>
        /// A "command-line" string (not a string array).
        /// </returns>
        public string sCommandLine()
        {
            foreach ( DictionaryEntry loEntry in this )
                if ( (null == loEntry.Value ? "" : loEntry.Value.ToString()).Contains(this.sNewLine) )
                    return this.sCommandBlock();

            StringBuilder lsbCommandLine = new StringBuilder();

            foreach ( DictionaryEntry loEntry in this )
            {
                string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());
                string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                if ( !lsValue.Contains(mcsSpcMark) && !lsValue.Contains(mcsArgMark) )
                {
                    lsbCommandLine.Append(mcsSpcMark + lsKey + mcsAsnMark + lsValue);
                }
                else
                {
                    int     liQteMark1Pos = lsValue.IndexOf(mccQteMark1);
                            if ( -1 == liQteMark1Pos ) liQteMark1Pos = lsValue.Length;
                    int     liQteMark2Pos = lsValue.IndexOf(mccQteMark2);
                            if ( -1 == liQteMark2Pos ) liQteMark2Pos = lsValue.Length;
                    string  lsQteMark = liQteMark1Pos < liQteMark2Pos ? mcsQteMark2 : mcsQteMark1;

                    lsbCommandLine.Append(mcsSpcMark + lsKey + mcsAsnMark + lsQteMark + lsValue + lsQteMark);
                }
            }

            return lsbCommandLine.ToString();
        }

        /// <summary>
        /// Returns the entire contents of the profile as a "command-line" string
        /// array (eg. <see langword='-Key1=one'/>, <see langword='-Key2=2'/>,
        /// <see langword='-Key3'/>, <see langword='-Key4="we have four"'/>).
        /// </summary>
        /// <returns>
        /// A "command-line" string array.
        /// </returns>
        public string[] sCommandLineArray()
        {
            string[] lsCommandLineArray = new String[this.Count];

            for ( int i = 0; i <= lsCommandLineArray.GetLength(0) - 1; i++ )
            {
                DictionaryEntry loEntry = (DictionaryEntry) base[i];

                lsCommandLineArray[i] = (null == loEntry.Key ? "" : loEntry.Key.ToString()) + mcsAsnMark + (null == loEntry.Value ? "" : loEntry.Value.ToString());
            }

            return lsCommandLineArray;
        }

        /// <summary>
        /// The profile entry key at the given zero-based index position.
        /// </summary>
        /// <param name="aiIndex">
        /// The zero based index into the list of profile entries.
        /// </param>
        /// <returns>
        /// The key string found at aiIndex.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <p>aiIndex is less than zero.</p>
        /// <p></p>
        /// -or-
        /// <p></p>
        /// aiIndex is equal to or greater than <see cref="ArrayList.Count"/>.
        /// </exception>
        public string sKey(int aiIndex)
        {
            return (null == oEntry(aiIndex).Key ? "" : oEntry(aiIndex).Key.ToString());
        }

        /// <summary>
        /// Returns the entire contents of the profile as an XML string. See
        /// the <see cref="sXmlXpath"/>, <see cref="sXmlKeyKey"/> and
        /// <see cref="sXmlValueKey"/> properties to learn what XML tags
        /// will be used.
        /// </summary>
        /// <param name="abStartDocument">
        /// Boolean determines if "start document" tags will be included in the
        /// XML returned.
        /// </param>
        /// <param name="abStandAlone">
        /// Boolean determines if the "stand alone" attribute will be included
        /// in the "start document" tag of the XML returned.
        /// </param>
        /// <returns>
        /// An XML string.
        /// </returns>
        public string sXml(bool abStartDocument, bool abStandAlone)
        {
            StringBuilder   lsbFileAsStream = new StringBuilder();
            StringWriter    loStringWriter = null;
            XmlTextWriter   loXmlTextWriter = null;

            try
            {
                loXmlTextWriter = new XmlTextWriter(loStringWriter = new StringWriter(lsbFileAsStream));
                loXmlTextWriter.Formatting = Formatting.Indented;

                if ( abStartDocument )
                {
                    if ( abStandAlone )
                    {
                        loXmlTextWriter.WriteStartDocument(true);
                    }
                    else
                    {
                        // Don't even bother to write a "standalone" attribute.
                        loXmlTextWriter.WriteStartDocument();
                    }
                }

                string[] lsXpathArray = this.sXmlXpath.Split('/');

                for ( int i = 0; i < lsXpathArray.Length; i++ )
                {
                    if ( i < lsXpathArray.Length - 1 )
                    {
                        loXmlTextWriter.WriteStartElement(lsXpathArray[i]);
                    }
                    else
                    {
                        // We use "lbSaveSansCmdLine" below instead of "mbSaveSansCmdLine" for the
                        // needed side effects. Also, we don't want "-SaveSansCmdLine" added here.
                        bool    lbRemoveSaveSansCmdLineKey = !this.ContainsKey("-SaveSansCmdLine");
                        bool    lbSaveSansCmdLine = this.bSaveSansCmdLine;
                                if ( lbRemoveSaveSansCmdLineKey )
                                    this.Remove("-SaveSansCmdLine");

                        Hashtable loAppendMatchingKeysMap = new Hashtable();

                        foreach ( DictionaryEntry loEntry in this )
                        {
                            string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());
                            string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                            this.AppendEntry(lsKey, lsValue, lbSaveSansCmdLine, loAppendMatchingKeysMap, loXmlTextWriter, lsXpathArray[i]);
                        }
                    }
                }

                for ( int i = 0; i < lsXpathArray.Length - 1; i++ )
                {
                    loXmlTextWriter.WriteEndElement();
                }

                if ( abStartDocument )
                    loXmlTextWriter.WriteEndDocument();

                // Replace entities since they have no impact on subsequent successful XML reads.
                lsbFileAsStream.Replace("&#xD;&#xA;", this.sNewLine);

                // Replace "utf-16" with "UTF-8" to allow current browser support.
                lsbFileAsStream.Replace("encoding=\"utf-16\"", "encoding=\"UTF-8\"");
            }
            finally
            {
                if ( null != loXmlTextWriter )
                    loXmlTextWriter.Close();
                if ( null != loStringWriter )
                    loStringWriter.Close();
            }

            return lsbFileAsStream.ToString();
        }

        /// <summary>
        /// Returns a subset of the profile as a new profile. All items
        /// that match asKey will be included.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find items in the profile. "*" or a regular
        /// expression may be included.
        /// </param>
        /// <returns>
        /// A new profile object containing the items found.
        /// </returns>
        public tvProfile oOneKeyProfile(string asKey)
        {
            return this.oOneKeyProfile(asKey, false);
        }

        /// <summary>
        /// Returns a subset of the profile as a new profile. All items
        /// that match asKey will be included.
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find items in the profile. "*" or a regular
        /// expression may be included.
        /// </param>
        /// <param name="abRemoveKeyPrefix">
        /// If true and asKey contains "*" or ".*", asKey (sans the wildcards)
        /// will be removed from each key prior to its addition to the new
        /// profile.
        /// </param>
        /// <returns>
        /// A new profile object containing the items found.
        /// </returns>
        public tvProfile oOneKeyProfile(string asKey, bool abRemoveKeyPrefix)
        {
            string  lsKeyPrefixToRemove = asKey.Replace(".*","").Replace("*","");
                    if ( asKey == lsKeyPrefixToRemove || String.IsNullOrEmpty(lsKeyPrefixToRemove) )
                    {
                        // If the given key contains no wildcards, it's not really a prefix.
                        abRemoveKeyPrefix = false;
                    }
            tvProfile loProfile = new tvProfile();

            if ( mbUseLiteralsOnly )
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());

                    if ( lsKey == asKey )
                    {
                        if ( abRemoveKeyPrefix )
                        {
                            loProfile.Add(mcsArgMark + lsKey.Replace(lsKeyPrefixToRemove, ""), loEntry.Value);
                        }
                        else
                        {
                            loProfile.Add(lsKey, loEntry.Value);
                        }
                    }
                }
            }
            else
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(lsKey, lsExpression, RegexOptions.IgnoreCase) )
                    {
                        if ( abRemoveKeyPrefix )
                        {
                            loProfile.Add(mcsArgMark + lsKey.Replace(lsKeyPrefixToRemove, ""), loEntry.Value);
                        }
                        else
                        {
                            loProfile.Add(lsKey, loEntry.Value);
                        }
                    }
                }
            }

            return loProfile;
        }

        /// <summary>
        /// Returns a subset of the profile as a trimmed string array.
        /// All items that match asKey are included. Note: only values are included
        /// (ie. no keys).
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find items in the profile. "*" or a regular
        /// expression may be included.
        /// </param>
        /// <returns>
        /// A string array containing the values found. Any environment variables
        /// embedded are expanded and each resulting string is trimmed of leading
        /// and trailing spaces.
        /// </returns>
        public string[] sOneKeyArray(string asKey)
        {
            StringBuilder lsbList = new StringBuilder();

            if ( mbUseLiteralsOnly )
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    if ( asKey == (null == loEntry.Key ? "" : loEntry.Key.ToString()) )
                    {
                        string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                        // The return limit of ExpandEnvironmentVariables is 32K. The difference allows for variable expansion.
                        if ( 32000 > lsValue.Length )
                        {
                            lsValue =  Environment.ExpandEnvironmentVariables(lsValue);
                        }

                        lsbList.Append(lsValue.Trim() + mccSplitMark);
                    }
                }
            }
            else
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(null == loEntry.Key ? "" : loEntry.Key.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                        // The return limit of ExpandEnvironmentVariables is 32K. The difference allows for variable expansion.
                        if ( 32000 > lsValue.Length )
                        {
                            lsValue =  Environment.ExpandEnvironmentVariables(lsValue);
                        }

                        lsbList.Append(lsValue.Trim() + mccSplitMark);
                    }
                }
            }

            string lsList = lsbList.ToString();

            if ( lsList.EndsWith(mcsSplitMark) )
            {
                return lsList.Remove(lsList.Length - 1, 1).Split(mccSplitMark);
            }
            else
            {
                return new String[0];
            }
        }

        /// <summary>
        /// Returns a subset of the profile as a string array (not trimmed). All
        /// items that match asKey are included. Note: only values are included
        /// (ie. no keys).
        /// </summary>
        /// <param name="asKey">
        /// The key string used to find items in the profile. "*" or a regular
        /// expression may be included.
        /// </param>
        /// <returns>
        /// A string array containing the values found. Any environment variables
        /// embedded are expanded and each resulting string is NOT trimmed of
        /// leading or trailing spaces.
        /// </returns>
        public string[] sOneKeyArrayNoTrim(string asKey)
        {
            StringBuilder lsbList = new StringBuilder();

            if ( mbUseLiteralsOnly )
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    if ( asKey == (null == loEntry.Key ? "" : loEntry.Key.ToString()) )
                    {
                        string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                        // The return limit of ExpandEnvironmentVariables is 32K. The difference allows for variable expansion.
                        if ( 32000 > lsValue.Length )
                        {
                            lsValue =  Environment.ExpandEnvironmentVariables(lsValue);
                        }

                        lsbList.Append(lsValue + mccSplitMark);
                    }
                }
            }
            else
            {
                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsExpression = this.sExpression(asKey);

                    if ( null != lsExpression && Regex.IsMatch(null == loEntry.Key ? "" : loEntry.Key.ToString(), lsExpression, RegexOptions.IgnoreCase) )
                    {
                        string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                        // The return limit of ExpandEnvironmentVariables is 32K. The difference allows for variable expansion.
                        if ( 32000 > lsValue.Length )
                        {
                            lsValue =  Environment.ExpandEnvironmentVariables(lsValue);
                        }

                        lsbList.Append(lsValue + mccSplitMark);
                    }
                }
            }

            string lsList = lsbList.ToString();

            if ( lsList.EndsWith(mcsSplitMark) )
            {
                return lsList.Remove(lsList.Length - 1, 1).Split(mccSplitMark);
            }
            else
            {
                return new String[0];
            }
        }

        /// <summary>
        /// Returns a full path/file string relative to the EXE file
        /// location, if asPathFile is only a filename. Otherwise, asPathFile is
        /// returned unchanged. This feature is useful for locating ancillary
        /// files in the same folder as the EXE file.
        /// </summary>
        /// <param name="asPathFile">
        /// A full path/file string or a filename only.
        /// </param>
        /// <returns>
        /// A full path/file string.
        /// </returns>
        public string sRelativeToExePathFile(string asPathFile)
        {
            string lsRelativeToExePathFile = null;

            if ( null != asPathFile )
            {
                if ( String.IsNullOrEmpty(Path.GetPathRoot(asPathFile)) )
                {
                    lsRelativeToExePathFile = Path.Combine(Path.GetDirectoryName(this.sExePathFile), asPathFile);
                }
                else
                {
                    lsRelativeToExePathFile = asPathFile;
                }
            }

            return lsRelativeToExePathFile;
        }

        /// <summary>
        /// Returns a full path/file string relative to the profile file
        /// location, if asPathFile is only a filename. Otherwise, asPathFile is
        /// returned unchanged. This feature is useful for locating ancillary
        /// files in the same folder as the profile file.
        /// </summary>
        /// <param name="asPathFile">
        /// A full path/file string or a filename only.
        /// </param>
        /// <returns>
        /// A full path/file string.
        /// </returns>
        public string sRelativeToProfilePathFile(string asPathFile)
        {
            string lsRelativeToProfilePathFile = null;

            if ( null != asPathFile )
            {
                if ( String.IsNullOrEmpty(this.sActualPathFile) )
                {
                    if ( String.IsNullOrEmpty(Path.GetPathRoot(asPathFile)) )
                    {
                        lsRelativeToProfilePathFile = Path.Combine(Path.GetDirectoryName(this.sDefaultPathFile), asPathFile);
                    }
                    else
                    {
                        lsRelativeToProfilePathFile = asPathFile;
                    }
                }
                else
                {
                    if ( String.IsNullOrEmpty(Path.GetPathRoot(asPathFile)) )
                    {
                        lsRelativeToProfilePathFile = Path.Combine(Path.GetDirectoryName(this.sActualPathFile), asPathFile);
                    }
                    else
                    {
                        lsRelativeToProfilePathFile = asPathFile;
                    }
                }
            }

            return lsRelativeToProfilePathFile;
        }

        /// <summary>
        /// Reloads the profile from the original text file used to load it and
        /// merges in the original command-line as well. Any changes to the
        /// profile in memory (not saved) since the last load will be lost.
        /// </summary>
        public void Reload()
        {
            this.Clear();

            if ( null != this.sLoadedPathFile )
                this.Load(this.sLoadedPathFile, tvProfileLoadActions.Append);

            this.LoadFromCommandLineArray(this.sInputCommandLineArray, tvProfileLoadActions.Merge);
        }

        /// <summary>
        /// Loads the profile with items from the given text file.
        /// <p>
        /// If <see cref="bUseXmlFiles"/> is true, text files will be read assuming
        /// standard XML "configuration file" format rather than line delimited
        /// "command-line" format.
        /// </p>
        /// </summary>
        /// <param name="asPathFile">
        /// The path/file location of the text file to load. This value will
        /// be used to set <see cref="sLoadedPathFile"/> after a successful load.
        /// </param>
        /// <param name="aeLoadAction">
        /// The action to take while loading profile items.
        /// See <see cref="tvProfileLoadActions"/>
        /// </param>
        public void Load(

                  string asPathFile
                , tvProfileLoadActions aeLoadAction
                )
        {
            // Check for the existence of one of several default filenames
            // (starting with asPathFile "as is"). Returned null means none exist.
            string  lsPathFile = this.sFileExistsFromList(asPathFile);
                    if ( String.IsNullOrEmpty(lsPathFile) )
                        lsPathFile = this.sFileExistsFromList(this.sRelativeToProfilePathFile(asPathFile));
            string  lsFilnameOnly = Path.GetFileNameWithoutExtension(this.sExePathFile);

            if ( String.IsNullOrEmpty(lsPathFile) )
            {
                if ( String.IsNullOrEmpty(asPathFile) )
                {
                    lsPathFile = this.sDefaultPathFile;
                }
                else
                {
                    lsPathFile = asPathFile;
                }

                // Pause to allow original instance (if any) to close.
                System.Threading.Thread.Sleep(200);

                // Count running instances.
                string      lsExeName = Path.GetFileNameWithoutExtension(this.sExePathFile);
                Process[]   loProcessesArray = Process.GetProcessesByName(lsExeName);

                if ( loProcessesArray.Length > 1 )
                {
                    if ( null != tvProfile.oMsgBoxShow2 )
                        tvProfile.oMsgBoxShow2.Invoke(null
                                , new object[]{String.Format("\"{0}\" is already running. Please close it and try again."
                                    , Path.GetFileName(this.sExePathFile)), lsExeName});

                    this.bExit = true;
                    return;
                }
                else
                if ( tvProfileFileCreateActions.PromptToCreateFile == this.eFileCreateAction )
                {
                    string  lsNewPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), lsFilnameOnly);
                    string  lsNewExePathFile = Path.Combine(lsNewPath, Path.GetFileName(this.sExePathFile));
                    bool    lbHasArgs = false;
                            if ( null != this.sInputCommandLineArray )
                                foreach ( string lsItem in this.sInputCommandLineArray )
                                {
                                    if ( lsItem.Trim().StartsWith(mcsArgMark) )
                                    {
                                        lbHasArgs = true;
                                        break;
                                    }
                                }

                    // If the EXE is not already in a folder with a matching name and if
                    // the EXE is not already installed in a typical installation folder, 
                    // and if there are no arguments passed on the command-line, proceed.
                    if ( !this.bInOwnFolder && !this.bInstalledAlready && !lbHasArgs )
                    {
                        bool    lbDoCopy = true;

                        if ( null != tvProfile.oMsgBoxShow4 )
                        {
                            string  lsMessage = String.Format(( Directory.Exists(lsNewPath)
                                    ? "an existing folder ({0}) on your desktop"
                                    : "a new folder ({0}) on your desktop" ), lsFilnameOnly);

                            lbDoCopy = /* DialogResult.OK */ 1 == (int)tvProfile.oMsgBoxShow4.Invoke(null, new object[]{String.Format(@"
For your convenience, this program will be copied
to {0}.

Depending on your system, this may take several seconds.  

Copy and proceed from there?

"
                                    , lsMessage), "Copy EXE to Desktop?", 1 /* MessageBoxButtons.OKCancel */, 32 /* MessageBoxIcon.Question */});
                        }
                        if ( lbDoCopy )
                        {
                            if ( !Directory.Exists(lsNewPath) )
                                Directory.CreateDirectory(lsNewPath);

                            File.Copy(this.sExePathFile, lsNewExePathFile, true);

                            ProcessStartInfo    loStartInfo = new ProcessStartInfo(lsNewExePathFile);
                                                loStartInfo.WorkingDirectory = Path.GetDirectoryName(lsNewExePathFile);
                            Process             loProcess = Process.Start(loStartInfo);
                        }

                        this.bExit = true;
                    }
                    else
                    if ( this.sExePathFile == lsNewExePathFile )
                    {
                        // The EXE has been moved to its own folder on the desktop.
                        // If it still exists on the desktop directly, remove it.

                        string lsOldExePathFile = Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Path.GetFileName(this.sExePathFile));

                        try
                        {
                            File.Delete(lsOldExePathFile);
                        }
                        catch
                        {
                            // Wait a moment ...
                            System.Threading.Thread.Sleep(200);

                            // Then try again.
                            try
                            {
                                File.Delete(lsOldExePathFile);
                            }
                            catch {}
                        }
                    }
                }

                // Although technically the file was not loaded (since it doesn't exist yet),
                // we consider it empty and we set the property to allow for subsequent saves.
                if ( !this.bExit )
                    this.sLoadedPathFile = lsPathFile;
            }
            else
            {
                string  lsFileAsStream = null;
                        this.UnlockProfileFile();

                        StreamReader loStreamReader = null;

                        try
                        {
                            loStreamReader = new StreamReader(lsPathFile);
                            lsFileAsStream = loStreamReader.ReadToEnd();
                            this.sLoadedPathFile = lsPathFile;
                        }
                        catch (IOException ex)
                        {
                            if ( ex.Message.Contains("being used by another process") )
                            {
                                // The app is most likely already running.
                                this.bExit = true;
                                return;
                            }
                            else
	                        {
                                // Wait a moment ...
                                System.Threading.Thread.Sleep(200);

                                // Then try again.
                                if ( null != loStreamReader )
                                    loStreamReader.Close();

                                loStreamReader = new StreamReader(lsPathFile);
                                lsFileAsStream = loStreamReader.ReadToEnd();
                                this.sLoadedPathFile = lsPathFile;
	                        }
                        }
                        finally
                        {
                            if ( null != loStreamReader )
                                loStreamReader.Close();
                        }

                        // If we can't lock the profile file,
                        // the app is most likely already running.
                        if ( !this.bLockProfileFile(lsPathFile) )
                        {
                            this.bExit = true;
                            return;
                        }

                int liDoOver = 1;

                do
                {
                    // mbUseXmlFiles is intentionally used here (instead of "this.bUseXmlFiles") to avoid side effects.
                    if ( mbUseXmlFiles )
                    {
                        try
                        {
                            this.LoadFromXml(lsFileAsStream, aeLoadAction);

                            // The profile file format is as expected. No do-over is needed.
                            liDoOver = 0;

                            if ( !this.bUseXmlFiles )
                            {
                                lsPathFile = this.sReformatProfileFile(lsPathFile);
                            }
                        }
                        catch
                        {
                            // mbUseXmlFiles is intentionally used here (instead of "this.bUseXmlFiles") to avoid side effects.
                            mbUseXmlFiles = false;
                        }
                    }

                    // mbUseXmlFiles is intentionally used here (instead of "this.bUseXmlFiles") to avoid side effects.
                    if ( !mbUseXmlFiles )
                    {
                        if ( lsFileAsStream.Length > 11 )
                        {
                            // Look for the XML tag only near the top (an XML document could be embedded way below).
                            if ( -1 == lsFileAsStream.IndexOf("<?xml versi", 0, 11) )
                            {
                                // The profile file format is as expected. No do-over is needed.
                                liDoOver = 0;

                                // The default file format is line delimited "command-line" format.
                                this.LoadFromCommandLineArray(this.sReplaceNewLine(lsFileAsStream).Split(mccSplitMark), aeLoadAction);

                                if ( this.bUseXmlFiles )
                                {
                                    lsPathFile = this.sReformatProfileFile(lsPathFile);
                                }
                            }
                            else
                            {
                                // mbUseXmlFiles is intentionally used here (instead of "this.bUseXmlFiles") to avoid side effects.
                                mbUseXmlFiles = true;
                                this.UnlockProfileFile();
                            }
                        }
                    }
                }
                while ( liDoOver-- > 0 );

                if ( !this.bLockProfileFile(lsPathFile) )
                    this.bExit = true;
            }

            // If it doesn't already exist, create the file.
            if ( !this.bExit
                    && !File.Exists(lsPathFile)
                    && tvProfileFileCreateActions.NoFileCreate != this.eFileCreateAction )
            {
                this.Save(lsPathFile);

                // In this case, consider the file loaded also.
                this.sLoadedPathFile = lsPathFile;
            }
        }

        /// <summary>
        /// Loads the profile with items from the given "command-line" string.
        /// </summary>
        /// <param name="asCommandLine">
        /// A string (not a string array) of the form:
        /// <see langword='-Key1=one -Key2=2 -Key3 -Key4="we have four"'/>.
        /// </param>
        /// <param name="aeLoadAction">
        /// The action to take while loading profile items.
        /// See <see cref="tvProfileLoadActions"/>
        /// </param>
        public void LoadFromCommandLine(

                  string asCommandLine
                , tvProfileLoadActions aeLoadAction
                )
        {
            if ( String.IsNullOrEmpty(asCommandLine) )
                return;

            // Remove any leading spaces or tabs so that mccSplitMark becomes the first char.
            asCommandLine = asCommandLine.TrimStart(mccSpcMark);
            asCommandLine = asCommandLine.TrimStart('\t');

            if ( -1 != asCommandLine.IndexOf('\r') || -1 != asCommandLine.IndexOf('\n') )
            {
                // If the command-line is actually already line delimited, then we're practically done.
                this.LoadFromCommandLineArray(this.sReplaceNewLine(asCommandLine).Split(mccSplitMark), aeLoadAction);
            }
            else
            {
                StringBuilder   lsbNewCommandLine = new StringBuilder();

                bool lbQteMark1On = false;
                bool lbQteMark2On = false;
                char lcCurrent = mccNulMark;

                for ( int i = 0; i <= asCommandLine.Length - 1; i++ )
                {
                    char    lcPrevious;
                            lcPrevious = lcCurrent;
                            lcCurrent = asCommandLine[i];

                    if ( mccQteMark1 == lcCurrent )
                    {
                        if ( lbQteMark1On )
                        {
                            lbQteMark1On = false;
                        }
                        else
                        if ( !lbQteMark2On )
                        {
                            lbQteMark1On = true;
                        }
                    }
                    else
                    if ( mccQteMark2 == lcCurrent )
                    {
                        if ( lbQteMark2On )
                        {
                            lbQteMark2On = false;
                        }
                        else
                        if ( !lbQteMark1On )
                        {
                            lbQteMark2On = true;
                        }
                    }

                    if ( lbQteMark1On || lbQteMark2On || mccAsnMark == lcPrevious || mccArgMark != lcCurrent )
                    {
                        // Quotes allow for nested argument lists.
                        // "|| mccAsnMark == lcPrevious" allows for easy negative numbers (ie. without an escape - leading whitespace not allowed).
                        lsbNewCommandLine.Append(lcCurrent);
                    }
                    else
                    {
                        // Eliminate whitespace between arguments as we go.
                        lsbNewCommandLine = new StringBuilder(lsbNewCommandLine.ToString().Trim());
                        lsbNewCommandLine.Append(mccSplitMark);
                        lsbNewCommandLine.Append(mcsArgMark);
                    }
                }

                //Cleanup any remaining tabs.
                lsbNewCommandLine.Replace('\t', mccSpcMark);

                //The first occurrence of the separator must be removed.
                this.LoadFromCommandLineArray(lsbNewCommandLine.ToString().TrimStart(mccSplitMark).Split(mccSplitMark), aeLoadAction);
            }
        }

        /// <summary>
        /// Loads the profile with items from the given "command-line"
        /// string array.
        /// </summary>
        /// <param name="asCommandLineArray">
        /// A string array of the form: <see langword='-Key1=one'/>, <see langword='-Key2=2'/>,
        /// <see langword='-Key3'/>, <see langword='-Key4="we have four"'/>.
        /// </param>
        /// <param name="aeLoadAction">
        /// The action to take while loading profile items.
        /// See <see cref="tvProfileLoadActions"/>
        /// </param>
        public void LoadFromCommandLineArray(

                  string[] asCommandLineArray
                , tvProfileLoadActions aeLoadAction
                )
        {
            if ( tvProfileLoadActions.Overwrite == aeLoadAction )
            {
                this.Clear();
            }

            if ( null != asCommandLineArray )
            {
                string      lsBlockKey = null;
                string      lsBlockValue = "";
                string      lsBlockEnd = null;
                string      lsBlockExc = null;
                Hashtable   loMergeKeysMap = new Hashtable();

                foreach ( string lsItem in asCommandLineArray )
                {
                    bool lbIsArg;

                    if ( !(lbIsArg = lsItem.TrimStart().StartsWith(mcsArgMark))
                            && (String.IsNullOrEmpty(lsBlockEnd) || !lsItem.Contains(lsBlockEnd)) )
                    {
                        if ( null != lsBlockKey )
                        {
                            lsBlockValue += lsItem + this.sNewLine;
                        }

                        // If an item does not start with a mcsArgMark
                        // and is not within a block, ignore it.
                    }
                    else
                    {
                        string  lsKey   = null;
                        string  lsValue = null;
                        object  loValue = null;
                        int     liPos   = 0;

                        if ( !lbIsArg )
                        {
                            // lsBlockEnd must be in lsItem.

                            liPos = lsItem.IndexOf(lsBlockEnd);
                            lsBlockValue += lsItem.Substring(0, liPos) + this.sNewLine;
                            lsKey = lsBlockKey;
                            lsValue = mcsBlockEndMark;

                            // This is the excess after the end of the block.
                            // This will be discarded (at least for now).
                            lsBlockExc = lsItem.Substring(liPos + lsBlockEnd.Length);
                        }
                        else
                        {
                            liPos = lsItem.IndexOf(mcsAsnMark);

                            if ( -1 == liPos )
                            {
                                lsKey = lsItem.Trim();
                                loValue = true;
                            }
                            else
                            {
                                bool lbQteMark1 = false;
                                bool lbQteMark2 = false;

                                lsKey = lsItem.Substring(0, liPos).Trim();
                                lsValue = lsItem.Substring(liPos + 1).Trim();

                                if ( lsValue.StartsWith(mcsQteMark1) && lsValue.EndsWith(mcsQteMark1) )
                                    lbQteMark1 = true;
                                else
                                if ( lsValue.StartsWith(mcsQteMark2) && lsValue.EndsWith(mcsQteMark2) )
                                    lbQteMark2 = true;

                                if ( !lbQteMark1 && !lbQteMark2 )
                                {
                                    // This is intentionally not trimmed.
                                    loValue = lsItem.Substring(liPos + 1);
                                }
                                else
                                {
                                    // First, remove quotation marks (if any).
                                    if ( lsValue.Length < 2 )
                                        loValue = "";
                                    else
                                        loValue = lsValue.Substring(1, lsValue.Length - 2);
                                }
                            }

                            lsValue = (null == loValue ? "" : loValue.ToString());
                        }

                        if ( null != lsBlockKey )
                        {
                            if ( mcsBlockEndMark == lsValue && lsBlockKey == lsKey )
                            {
                                lsBlockKey = null;
                                loValue = lsBlockValue;
                                lsBlockEnd = null;
                                lsBlockExc = null;
                            }
                            else
                            {
                                lsBlockValue += lsItem + this.sNewLine;
                            }
                        }
                        else if ( mcsBlockBegMark == lsValue )
                        {
                            lsBlockKey = lsKey;
                            lsBlockValue = "";
                            lsBlockEnd = lsBlockKey + mcsAsnMark + mcsBlockEndMark;
                        }

                        if ( String.IsNullOrEmpty(lsBlockKey) )
                        {
                            switch ( aeLoadAction )
                            {
                                case tvProfileLoadActions.Append:
                                case tvProfileLoadActions.Overwrite:

                                    this.Add(lsKey, loValue);
                                    break;

                                case tvProfileLoadActions.Merge:

                                    bool    lbDiscard = this.bSaveSansCmdLine;  // We need only the side-effects.
                                    int     liIndex = this.IndexOfKey(lsKey);

                                    // Replace wildcard keys with the first key match, if any.
                                    lsKey = ( -1 == liIndex ? lsKey : this.sKey(liIndex) );

                                    if ( loMergeKeysMap.ContainsKey(lsKey) )
                                    {
                                        // Set the search index to force adding this key.
                                        liIndex = -1;
                                    }
                                    else
                                    {
                                        if ( -1 != liIndex )
                                        {
                                            if ( 1 == this.iKeyCount(lsKey) )
                                            {
                                                // Match original singleton's position.
                                                this[liIndex] = loValue;
                                            }
                                            else
                                            {
                                                // Remove all previous entries with this key (presumably from a file).
                                                this.Remove(lsKey);

                                                // Set the search index to force adding this key with its overriding value.
                                                liIndex = -1;
                                            }
                                        }

                                        // Add to the merge key map to prevent any further removals of this key.
                                        loMergeKeysMap.Add(lsKey, null);
                                    }

                                    if ( -1 == liIndex )
                                    {
                                        // Don't add keys that contain '*'.
                                        if ( -1 == lsKey.IndexOf('*') )
                                            this.Add(lsKey, loValue);
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Loads the profile with items from the given XML string. See the
        /// <see cref="sXmlXpath"/>, <see cref="sXmlKeyKey"/> and
        /// <see cref="sXmlValueKey"/> properties to learn what XML tags are
        /// expected.
        /// </summary>
        /// <param name="asXml">
        /// An XML document as a string.
        /// </param>
        /// <param name="aeLoadAction">
        /// The action to take while loading profile items.
        /// See <see cref="tvProfileLoadActions"/>
        /// </param>
        public void LoadFromXml(

                  string asXml
                , tvProfileLoadActions aeLoadAction
                )
        {
            if ( tvProfileLoadActions.Overwrite == aeLoadAction )
            {
                this.Clear();
            }

            XmlDocument loXmlDoc = new XmlDocument();
                        loXmlDoc.LoadXml(asXml);

            foreach ( XmlNode loNode in loXmlDoc.SelectNodes(this.sXmlXpath) )
            {
                string lsKey = loNode.Attributes[this.sXmlKeyKey].Value;
                string lsValue = loNode.Attributes[this.sXmlValueKey].Value.StartsWith(this.sNewLine)
                        ? loNode.Attributes[this.sXmlValueKey].Value.Substring(this.sNewLine.Length
                                , loNode.Attributes[this.sXmlValueKey].Value.Length - 2 * this.sNewLine.Length)
                        : loNode.Attributes[this.sXmlValueKey].Value;

                switch ( aeLoadAction )
                {
                    case tvProfileLoadActions.Append:
                    case tvProfileLoadActions.Overwrite:

                        this.Add(lsKey, lsValue);
                        break;

                    case tvProfileLoadActions.Merge:

                        this.bSaveEnabled = false;

                        int liIndex = this.IndexOfKey(lsKey);

                        if ( -1 != liIndex )
                        {
                            // Set each entry with a matching key to the given value.

                            foreach ( DictionaryEntry loEntry in this.oOneKeyProfile(lsKey, false) )
                            {
                                this.SetByIndex(this.IndexOfKey(null == loEntry.Key ? "" : loEntry.Key.ToString()), lsValue);
                            }
                        }
                        else
                        {
                            // Don't add keys that contain '*'.
                            if ( -1 == lsKey.IndexOf('*') )
                                this.Add(lsKey, lsValue);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Saves the contents of the profile as a text file.
        /// </summary>
        /// <param name="asPathFile">
        /// The path/file location to save the profile file to. This value will
        /// be used to set <see cref="sActualPathFile"/> after a successful save.
        /// </param>
        public void Save(string asPathFile)
        {
            this.sActualPathFile = asPathFile;
            this.Save();
        }

        /// <summary>
        /// <p>
        /// Saves the contents of the profile as a text file using the location
        /// referenced in <see cref="sActualPathFile"/>. <see cref="sActualPathFile"/>
        /// will have the same value as <see cref="sLoadedPathFile"/> after a successful
        /// load from a text file.
        /// </p>
        /// <p>
        /// If <see cref="bUseXmlFiles"/> is true, text files will be written in
        /// standard XML "configuration file" format rather than line delimited
        /// "command-line" format.
        /// </p>
        /// </summary>
        public void Save()
        {
            if ( !this.bSaveEnabled || String.IsNullOrEmpty(this.sActualPathFile) )
            {
                return;
            }

            bool    lbAlreadyThere = File.Exists(this.sActualPathFile);
            string  lsFileAsStream = null;

            if ( this.bUseXmlFiles )
            {
                if ( !lbAlreadyThere )
                {
                    lsFileAsStream = this.sXml(true, false) + this.sNewLine;
                }
                else
                {
                    string      lsXmlXpath = this.sXmlXpath;
                    XmlDocument loXmlDocument = new XmlDocument();
                                try
                                {
                                    this.UnlockProfileFile();
                                    loXmlDocument.Load(this.sActualPathFile);
                                }
                                catch (XmlException)
                                {
                                    tvProfile   loProfie = new tvProfile(this.sCommandLine());
                                                this.Load(this.sActualPathFile, tvProfileLoadActions.Overwrite);
                                                this.LoadFromCommandLine(loProfie.sCommandLine(), tvProfileLoadActions.Merge);

                                    lsFileAsStream = this.sXml(true, false) + this.sNewLine;
                                }

                    if ( null == lsFileAsStream )
                    {
                        XmlNode     loXmlNode = loXmlDocument.SelectSingleNode("configuration/appSettings");
                                    if ( null != loXmlNode )
                                    {
                                        // Replace all application settings already there.
                                        this.sXmlXpath = "add";
                                        loXmlNode.InnerXml = this.sXml(false, false);
                                        this.sXmlXpath = lsXmlXpath;
                                    }
                                    else
                                    {
                                        loXmlNode = loXmlDocument.SelectSingleNode("configuration");
                                        if ( null == loXmlNode )
                                        {
                                            throw new Exception("XML configuration tags missing. Can't continue.");
                                        }
                                        else
                                        {
                                            // Add an application settings section.
                                            this.sXmlXpath = "appSettings/add";
                                            XmlDocumentFragment loXmlDocumentFragment = loXmlDocument.CreateDocumentFragment();
                                                                loXmlDocumentFragment.InnerXml = this.sXml(false, false);

                                            loXmlDocument.DocumentElement.InsertBefore(loXmlDocumentFragment, loXmlDocument.DocumentElement.FirstChild);
                                            this.sXmlXpath = lsXmlXpath;
                                        }
                                    }

                        // Replace entities since they have no impact on subsequent successful XML reads.
                        lsFileAsStream = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + this.sNewLine + XDocument.Parse(loXmlDocument.InnerXml).ToString().Replace("&#xD;&#xA;", this.sNewLine) + this.sNewLine;
                    }
                }
            }
            else
            {
                StringBuilder lsbFileAsStream = new StringBuilder(Path.GetFileName(this.sExePathFile) + this.sNewLine + this.sNewLine);

                // We use "lbSaveSansCmdLine" below instead of "mbSaveSansCmdLine" for the
                // needed side effects. Also, we don't want "-SaveSansCmdLine" added here.
                bool    lbRemoveSaveSansCmdLineKey = !this.ContainsKey("-SaveSansCmdLine");
                bool    lbSaveSansCmdLine = this.bSaveSansCmdLine;
                        if ( lbRemoveSaveSansCmdLineKey )
                            this.Remove("-SaveSansCmdLine");

                Hashtable loAppendMatchingKeysMap = new Hashtable();

                foreach ( DictionaryEntry loEntry in this )
                {
                    string lsKey = (null == loEntry.Key ? "" : loEntry.Key.ToString());
                    string lsValue = (null == loEntry.Value ? "" : loEntry.Value.ToString());

                    this.AppendEntry(lsKey, lsValue, lbSaveSansCmdLine, loAppendMatchingKeysMap, lsbFileAsStream);
                }

                lsFileAsStream = lsbFileAsStream.ToString();
            }

            this.UnlockProfileFile();

            string  lsPath = Path.GetDirectoryName(this.sActualPathFile);
                    if ( !String.IsNullOrEmpty(lsPath) )
                        Directory.CreateDirectory(lsPath);

            try
            {
                File.WriteAllText(this.sActualPathFile, lsFileAsStream);
            }
            catch (Exception)
            {
                // Wait a moment ...
                if ( null != tvProfile.oAppDoEvents )
                    tvProfile.oAppDoEvents.Invoke(null, null);
                System.Threading.Thread.Sleep(200);

                // Then try again.
                File.WriteAllText(this.sActualPathFile, lsFileAsStream);
            }

            this.bLockProfileFile(this.sActualPathFile);

            if ( !lbAlreadyThere )
                this.bFileJustCreated = true;
        }

        /// <summary>
        /// A custom enumerator. This is necessary since the indexer of this class
        /// returns the object found by index or by key rather than the underlying
        /// DictionaryEntry object that contains the "key/value" pair.
        /// </summary>
        public override IEnumerator GetEnumerator()
        {
            return new tvProfileEnumerator(this);
        }

        /// <summary>
        /// This exception is thrown by <see langword='Add(object)'/> if the given object
        /// is anything other than a DictionaryEntry.
        /// </summary>
        public class InvalidAddType : Exception
        {
            /// <summary>
            /// Returns "Only DictionaryEntry objects may be added with this method."
            /// </summary>
            public override string Message
            {
                get
                {
                    return "Only DictionaryEntry objects may be added with this method.";
                }
            }

        }

        #region "Private Members"

        private bool bLockProfileFile(string asPathFile)
        {
            bool lbLockProfileFile = false;

            if ( null != moFileStreamProfileFileLock || !this.bEnableFileLock )
            {
                lbLockProfileFile = true;
            }
            else
            {
                try
                {
                    // mbUseXmlFiles is intentionally used here (instead of "this.bUseXmlFiles") to avoid side effects.
                    if ( !mbUseXmlFiles )
                        moFileStreamProfileFileLock =
                                File.Open(asPathFile, FileMode.Open, FileAccess.Read, FileShare.None);

                    lbLockProfileFile = true;
                }
                catch {/* Most likely trying to run more than one instance. Let main app handle it. */}
            }

            return lbLockProfileFile;
        }

        private string sExpression(string asSource)
        {
            if ( String.IsNullOrEmpty(asSource) )
                return null;

            string lsExpression = asSource;

            if ( -1 == lsExpression.IndexOf(".*") )
                lsExpression = lsExpression.Replace("*", ".*");

            if ( -1 == lsExpression.IndexOf("$") )
                lsExpression = lsExpression + "$";

            return lsExpression;
        }

        private string sFileExistsFromList(string asPathFile)
        {
            if ( null != asPathFile )
            {
                if ( File.Exists(asPathFile) )
                {
                    return asPathFile;
                }
            }
            else
            {
                string lsDefaultPathFileNoExt = Path.Combine(Path.GetDirectoryName(this.sDefaultPathFile)
                                                        , Path.GetFileNameWithoutExtension(this.sDefaultPathFile));

                foreach ( string lsItem in msDefaultFileExtArray )
                {
                    string lsDefaultPathFile = lsDefaultPathFileNoExt + lsItem;

                    if ( File.Exists(lsDefaultPathFile) )
                        return lsDefaultPathFile;
                }
            }

            return null;
        }

        private string sReformatProfileFile(string asPathFile)
        {
            if ( null != asPathFile)
            {
                if ( this.bDefaultFileReplaced )
                {
                    // Reuse the replacement pathfile.
                    string  lsTmpPathfile = this.sRelativeToProfilePathFile(Guid.NewGuid().ToString());
                            this.Save(lsTmpPathfile);           // This allows for an XmlDocument.Load()
                            File.Delete(this.sLoadedPathFile);  // of the original file during the save.
                            File.Move(lsTmpPathfile, this.sLoadedPathFile);

                    this.sActualPathFile = this.sLoadedPathFile;
                }
                else
                {
                    // Use a new pathfile (perhaps).
                    asPathFile = this.sDefaultPathFile;
                    this.Save(asPathFile);

                    if ( this.bSaveEnabled && this.sLoadedPathFile != asPathFile )
                    {
                        File.Delete(this.sLoadedPathFile);
                        this.sLoadedPathFile = asPathFile;
                    }
                }
            }

            return asPathFile;
        }

        private string sReplaceNewLine(string asSource)
        {
            if ( asSource.Contains(Environment.NewLine) )
                this.sNewLine = Environment.NewLine;    // this env
            else
            if ( -1 != asSource.IndexOf('\r') )
                this.sNewLine = "\r";                   // MacOS
            else
            if ( -1 != asSource.IndexOf('\n') )
                this.sNewLine = "\n";                   // *nix
            else
                return asSource;                        // no NL

            StringBuilder   lsbReplaceNewLine = new StringBuilder(asSource);
                            lsbReplaceNewLine = lsbReplaceNewLine.Replace(this.sNewLine, mcsSplitMark);

            return lsbReplaceNewLine.ToString();
        }

        private void AppendEntry(string asKey, string asValue, bool abSaveSansCmdLine, Hashtable aoAppendMatchingKeysMap, StringBuilder asbFileAsStream)
        {
            // "lbSaveSansCmdLine" is referenced here (in lieu of "this.bSaveSansCmdLine") to gain a little speed.
            if ( !abSaveSansCmdLine || null == moInputCommandLineProfile )
            {
                this.AppendEntry2(asKey, asValue, asbFileAsStream);
            }
            else
            if ( abSaveSansCmdLine )
            {
                if ( !moInputCommandLineProfile.ContainsKey(asKey) )
                {
                    this.AppendEntry2(asKey, asValue, asbFileAsStream);
                }
                else
                if ( !aoAppendMatchingKeysMap.ContainsKey(asKey) )
                {
                    // Add to the keys map to prevent any further appends of this matching key.
                    aoAppendMatchingKeysMap.Add(asKey, null);

                    foreach ( DictionaryEntry loEntry in moMatchCommandLineProfile.oOneKeyProfile(asKey) )
                    {
                        this.AppendEntry2(asKey, loEntry.Value.ToString(), asbFileAsStream);
                    }
                }
            }
        }

        private void AppendEntry2(string asKey, string asValue, StringBuilder asbFileAsStream)
        {
            if ( -1 == asValue.IndexOf(this.sNewLine) )
            {
                if ( -1 == asValue.IndexOf(mcsSpcMark) )
                    asbFileAsStream.Append(asKey + mcsAsnMark + asValue + this.sNewLine);
                else
                    asbFileAsStream.Append(asKey + mcsAsnMark + mcsQteMark1 + asValue + mcsQteMark1 + this.sNewLine);
            }
            else
            {
                asbFileAsStream.Append(asKey + mcsAsnMark + mcsBlockBegMark + this.sNewLine);
                asbFileAsStream.Append(asValue + ((asValue.EndsWith(this.sNewLine)) ? "" : this.sNewLine).ToString());
                asbFileAsStream.Append(asKey + mcsAsnMark + mcsBlockEndMark + this.sNewLine);
            }
        }

        private void AppendEntry(string asKey, string asValue, bool abSaveSansCmdLine, Hashtable aoAppendMatchingKeysMap, XmlTextWriter aoXmlTextWriter, string asXpathArrayItem)
        {
            // "lbSaveSansCmdLine" is referenced here (in lieu of "this.bSaveSansCmdLine") to gain a little speed.
            if ( !abSaveSansCmdLine || null == moInputCommandLineProfile )
            {
                this.AppendEntry2(asKey, asValue, aoXmlTextWriter, asXpathArrayItem);
            }
            else
            if ( abSaveSansCmdLine )
            {
                if ( !moInputCommandLineProfile.ContainsKey(asKey) )
                {
                    this.AppendEntry2(asKey, asValue, aoXmlTextWriter, asXpathArrayItem);
                }
                else
                if ( !aoAppendMatchingKeysMap.ContainsKey(asKey) )
                {
                    // Add to the keys map to prevent any further appends of this matching key.
                    aoAppendMatchingKeysMap.Add(asKey, null);

                    foreach ( DictionaryEntry loEntry in moMatchCommandLineProfile.oOneKeyProfile(asKey) )
                    {
                        this.AppendEntry2(asKey, loEntry.Value.ToString(), aoXmlTextWriter, asXpathArrayItem);
                    }
                }
            }
        }

        private void AppendEntry2(string asKey, string asValue, XmlTextWriter aoXmlTextWriter, string asXpathArrayItem)
        {
            aoXmlTextWriter.WriteStartElement(asXpathArrayItem);

                bool lbTextBlock = -1 != asValue.IndexOf(this.sNewLine);

                aoXmlTextWriter.WriteAttributeString(this.sXmlKeyKey, asKey);

                if ( lbTextBlock )
                {
                    aoXmlTextWriter.WriteAttributeString(this.sXmlValueKey, this.sNewLine + asValue + this.sNewLine);
                }
                else
                {
                    aoXmlTextWriter.WriteAttributeString(this.sXmlValueKey, asValue);
                }

            aoXmlTextWriter.WriteEndElement();
        }

        private void ReplaceDefaultProfileFromCommandLine(string[] asCommandLineArray)
        {
            tvProfile   loCommandLine = new tvProfile();
                        loCommandLine.LoadFromCommandLineArray(asCommandLineArray, tvProfileLoadActions.Append);
            string[]    lsIniKeys = { "-ini", "-ProfileFile" };
            int         liIniKeyIndex = - 1;
                        if ( loCommandLine.ContainsKey(lsIniKeys[0]) )
                        {
                            liIniKeyIndex = 0;
                        }
                        else
                        if ( loCommandLine.ContainsKey(lsIniKeys[1]) )
                        {
                            liIniKeyIndex = 1;
                        }
            string      lsProfilePathFile = null;
                        if ( -1 != liIniKeyIndex )
                        {
                            lsProfilePathFile = loCommandLine.sValue(lsIniKeys[liIniKeyIndex], "");
                        }
            bool        lbFirstArgIsFile = false;
            string      lsFirstArg = null;
                        try
                        {
                            if ( -1 != asCommandLineArray[0].IndexOf(".vshost.")
                                    || this.sRelativeToExePathFile(asCommandLineArray[0]) == this.sExePathFile )
                            {
                                lsFirstArg = asCommandLineArray[1];
                            }
                            else
                            {
                                lsFirstArg = asCommandLineArray[0];
                            }

                            if ( null != lsFirstArg
                                    && (File.Exists(lsFirstArg) || File.Exists(this.sRelativeToExePathFile(lsFirstArg))) )
                            {
                                if ( null != lsProfilePathFile )
                                {
                                    // If the first argument passed on the command-line is actually
                                    // a file (that exists) and if an -ini key was also provided, then
                                    // add the file reference to the profile using the "-File" key.
                                    lbFirstArgIsFile = true;
                                }
                                else
                                {
                                    // If no -ini key was passed, then assume the referenced file is
                                    // actually a profile file to be loaded.
                                    lsProfilePathFile = lsFirstArg;
                                }
                            }
                        }
                        catch {}

            if ( null != lsProfilePathFile )
            {
                // Load the referenced profile file.
                this.bDefaultFileReplaced = true;
                this.Load(lsProfilePathFile, tvProfileLoadActions.Overwrite);

                if ( !this.bExit )
                {
                    // We now need a slightly adjusted version of the given command-line
                    // (ie. sans the -ini key but with a -File key added, if appropriate).
                    if ( -1 != liIniKeyIndex )
                        loCommandLine.Remove(lsIniKeys[liIniKeyIndex]);
                    if ( lbFirstArgIsFile )
                        loCommandLine.Add("-File", lsFirstArg);

                    // Now merge in the original (adjusted) command-line
                    // (command-line items take precedence over file items).
                    this.LoadFromCommandLineArray(loCommandLine.sCommandLineArray(), tvProfileLoadActions.Merge);
                }
            }
        }

        private void UnlockProfileFile()
        {
            if ( null != moFileStreamProfileFileLock )
            {
                moFileStreamProfileFileLock.Close();
                moFileStreamProfileFileLock = null;
                GC.Collect();
            }
        }

        private static MethodInfo oAppDoEvents
        {
            get
            {
                if ( null == moAppDoEvents )
                {
                    // Setup DoEvents calls without the need for a compile time reference.
                    try
                    {
                        Assembly    loWinFormAssm = Assembly.Load(mcsWinFormsAssm);
                        Type        loObjectType = loWinFormAssm.GetType(mcsWinFormsAppType);
                                    moAppDoEvents = loObjectType.GetMethod("DoEvents");
                    }
                    catch {}
                }

                return moAppDoEvents;
            }
        }
        private static MethodInfo moAppDoEvents = null;

        private static MethodInfo oMsgBoxShow2
        {
            get
            {
                if ( null == moMsgBoxShow2 )
                {
                    // Setup MsgBox Show(string, string) calls without the need for a compile time reference.
                    try
                    {
                        Assembly    loWinFormAssm = Assembly.Load(mcsWinFormsAssm);
                        Type        loObjectType = loWinFormAssm.GetType(mcsWinFormsMsgBoxType);
                                    moMsgBoxShow2 = loObjectType.GetMethods()[13];
                    }
                    catch {}
                }

                return moMsgBoxShow2;
            }
        }
        private static MethodInfo moMsgBoxShow2 = null;

        private static MethodInfo oMsgBoxShow3
        {
            get
            {
                if ( null == moMsgBoxShow3 )
                {
                    // Setup MsgBox Show(string, string, enum) calls without the need for a compile time reference.
                    try
                    {
                        Assembly    loWinFormAssm = Assembly.Load(mcsWinFormsAssm);
                        Type        loObjectType = loWinFormAssm.GetType(mcsWinFormsMsgBoxType);
                                    moMsgBoxShow3 = loObjectType.GetMethods()[12];
                    }
                    catch {}
                }

                return moMsgBoxShow3;
            }
        }
        private static MethodInfo moMsgBoxShow3 = null;

        private static MethodInfo oMsgBoxShow4
        {
            get
            {
                if ( null == moMsgBoxShow4 )
                {
                    // Setup MsgBox Show(string, string, enum, enum) calls without the need for a compile time reference.
                    try
                    {
                        Assembly    loWinFormAssm = Assembly.Load(mcsWinFormsAssm);
                        Type        loObjectType = loWinFormAssm.GetType(mcsWinFormsMsgBoxType);
                                    moMsgBoxShow4 = loObjectType.GetMethods()[11];
                    }
                    catch {}
                }

                return moMsgBoxShow4;
            }
        }
        private static MethodInfo moMsgBoxShow4 = null;


        private const string        mcsLoadSaveDefaultExtension = ".txt";
        private const char          mccArgMark      = '-';
        private const string        mcsArgMark      = "-";
        private const char          mccAsnMark      = '=';
        private const string        mcsAsnMark      = "=";
        private const string        mcsBlockBegMark = "[";
        private const string        mcsBlockEndMark = "]";
        private const char          mccNulMark      = '\u0000';
        private string              mcsNulMark      = '\u0000'.ToString();
        private const char          mccQteMark1     = '\"';
        private const string        mcsQteMark1     = "\"";
        private const char          mccQteMark2     = '\'';
        private const string        mcsQteMark2     = "'";
        private const char          mccSpcMark      = ' ';
        private const string        mcsSpcMark      = " ";
        private const char          mccSplitMark    = '\u0001';
        private string              mcsSplitMark    = '\u0001'.ToString();
        private string              mcsSaveKey      = "-SaveProfile";
        private FileStream          moFileStreamProfileFileLock;
        private tvProfile           moInputCommandLineProfile;
        private tvProfile           moMatchCommandLineProfile;
        private static int          mciIntSizeInBytes = 4;
        private static string       mcsWinFormsAssm         = "System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
        private static string       mcsWinFormsAppType      = "System.Windows.Forms.Application";
        private static string       mcsWinFormsMsgBoxType   = "System.Windows.Forms.MessageBox";


        private class tvProfileKeyComparer : IComparer
        {
            public int Compare(object aoX, object aoY)
            {
                return ((DictionaryEntry)aoX).Key.ToString().CompareTo(((DictionaryEntry)aoY).Key.ToString());
            }
        }

        private class tvProfileStringValueComparer : IComparer
        {
            public int Compare(object aoX, object aoY)
            {
                return ((DictionaryEntry)aoX).Value.ToString().CompareTo(((DictionaryEntry)aoY).Value.ToString());
            }
        }

        private class tvProfileEnumerator : IEnumerator
        {
            int miIndex = -1;
            tvProfile moProfile;

            private tvProfileEnumerator(){}

            public tvProfileEnumerator(tvProfile aoProfile)
            {
                moProfile = aoProfile;
            }

            #region IEnumerator Members

            public void Reset()
            {
                miIndex = -1;
            }

            public object Current
            {
                get
                {
                    return moProfile.oEntry(miIndex);
                }
            }

            public bool MoveNext()
            {
                miIndex++;
                return miIndex < moProfile.Count;
            }

            #endregion
        }
        #endregion
    }
}
