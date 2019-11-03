using System;
using System.Windows;
using System.Windows.Input;
using tvToolbox;

public enum tvMessageBoxIcons
{
     Alert
    ,Done
    ,Error
    ,Exclamation
    ,Information
    ,None
    ,Ok
    ,Question
    ,Stop
    ,Warning
}

public enum tvMessageBoxButtons
{
     OK
    ,OKCancel
    ,YesNo
    ,YesNoCancel
}

public enum tvMessageBoxCheckBoxTypes
{
     DontAsk
    ,None
    ,SkipThis
}

public enum tvMessageBoxResults
{
     Cancel
    ,No
    ,None
    ,OK
    ,Yes
}

/// <summary>
/// Interaction logic for tvMessageBox.xaml
/// </summary>
public partial class tvMessageBox : Window
{
    public tvMessageBoxResults eTvMessageBoxResult = tvMessageBoxResults.None;

    private const string msProfilePromptKeyPrefix = "-MsgBoxPrompt";
    private const string msProfilePromptKeySuffix = "Answer";

    private bool mbIsShowWait = false;


    public tvMessageBox()
    {
        InitializeComponent();
    }


    public bool bDialogAccepted
    {
        get
        {
            return mbDialogAccepted;
        }
        set
        {
            mbDialogAccepted = value;
        }
    }
    private bool mbDialogAccepted = false;


    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        // This kludge is necessary to center in screen
        // since the dimensions were wrong to start with.
        this.Top = (System.Windows.SystemParameters.PrimaryScreenHeight - this.FirstBorder.ActualHeight) / 2;
        this.Left = (System.Windows.SystemParameters.PrimaryScreenWidth - this.FirstBorder.ActualWidth) / 2;

        // This kludge is necessary to have
        // the drag rectangle sized correctly.
        this.Height = this.FirstBorder.ActualHeight;
        this.Width = this.FirstBorder.ActualWidth;
    }

    private void btnOK_Click(object sender, RoutedEventArgs e)
    {
        this.eTvMessageBoxResult = tvMessageBoxResults.OK;
        this.bDialogAccepted = true;
        this.Close();
    }

    private void btnYes_Click(object sender, RoutedEventArgs e)
    {
        this.eTvMessageBoxResult = tvMessageBoxResults.Yes;
        this.Close();
    }

    private void btnNo_Click(object sender, RoutedEventArgs e)
    {
        this.eTvMessageBoxResult = tvMessageBoxResults.No;
        this.Close();
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
        this.eTvMessageBoxResult = tvMessageBoxResults.Cancel;
        this.Close();
    }

    private void Window_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if ( mbIsShowWait )
            this.Close();
        else
        if ( MouseButton.Left == e.ChangedButton )
            this.DragMove();
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        switch ( e.Key )
        {
            case Key.Enter:
                this.bDialogAccepted = true;
                break;
        }
    }

    private void SelectIcon(tvMessageBoxIcons aeTvMessageBoxIcon)
    {
        switch ( aeTvMessageBoxIcon )
        {
            case tvMessageBoxIcons.Alert:
                this.AlertIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Done:
                this.OkIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Error:
                this.ErrorIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Exclamation:
                this.DefaultIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Question:
                this.HelpIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Information:
                this.InfoIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.None:
                break;
            case tvMessageBoxIcons.Ok:
                this.OkIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Stop:
                this.ErrorIcon.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxIcons.Warning:
                this.AlertIcon.Visibility = Visibility.Visible;
                break;
        }
    }

    private void SelectButtons(tvMessageBoxButtons aeTvMessageBoxButton)
    {
        switch ( aeTvMessageBoxButton )
        {
            case tvMessageBoxButtons.OK:
                this.btnOK.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxButtons.OKCancel:
                this.btnOK.Visibility = Visibility.Visible;
                this.btnCancel.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxButtons.YesNo:
                this.btnYes.Visibility = Visibility.Visible;
                this.btnNo.Visibility = Visibility.Visible;
                break;
            case tvMessageBoxButtons.YesNoCancel:
                this.btnYes.Visibility = Visibility.Visible;
                this.btnNo.Visibility = Visibility.Visible;
                this.btnCancel.Visibility = Visibility.Visible;
                break;
        }
    }

    public static tvMessageBoxResults Show(Window aoWindow, string asMessageText)
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , null
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Information
                , false
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Information, false
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , null
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , false
                );
    }

    public static tvMessageBoxResults ShowModeless(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , true
                );
    }

    public static tvMessageBoxResults ShowModeless(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , tvMessageBoxCheckBoxTypes aeTvMessageBoxCheckBoxType
            , tvProfile aoProfile
            , string asProfilePromptKey
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , true
                , aeTvMessageBoxCheckBoxType
                , aoProfile
                , asProfilePromptKey
                , tvMessageBoxResults.None
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , tvMessageBoxCheckBoxTypes aeTvMessageBoxCheckBoxType
            , tvProfile aoProfile
            , string asProfilePromptKey
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , false
                , aeTvMessageBoxCheckBoxType
                , aoProfile
                , asProfilePromptKey
                , tvMessageBoxResults.None
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , false
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , bool abShowModeless
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , abShowModeless
                , tvMessageBoxCheckBoxTypes.None
                , null
                , null
                , tvMessageBoxResults.None
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , tvMessageBoxCheckBoxTypes aeTvMessageBoxCheckBoxType
            , tvProfile aoProfile
            , string asProfilePromptKey
            , tvMessageBoxResults aeTvMessageBoxResultsOverride
            )
    {
        return tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , aeTvMessageBoxButtons
                , aeTvMessageBoxIcon
                , false
                , aeTvMessageBoxCheckBoxType
                , aoProfile
                , asProfilePromptKey
                , aeTvMessageBoxResultsOverride
                );
    }

    public static tvMessageBoxResults Show(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxButtons aeTvMessageBoxButtons
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , bool abShowModeless
            , tvMessageBoxCheckBoxTypes aeTvMessageBoxCheckBoxType
            , tvProfile aoProfile
            , string asProfilePromptKey
            , tvMessageBoxResults aeTvMessageBoxResultsOverride
            )
    {
        tvMessageBoxResults liTvMessageBoxResult = tvMessageBoxResults.None;

        string  lsPromptAnswerKey = null;
        bool    lbUseCheckBox = tvMessageBoxCheckBoxTypes.None != aeTvMessageBoxCheckBoxType;
                if ( lbUseCheckBox )
                {
                    // Insert the prompt key prefix if it's not already there. A common prefix
                    // is necessary to allow for the removal of all prompt keys as needed.
                    if ( !asProfilePromptKey.StartsWith(msProfilePromptKeyPrefix) )
                    {
                        // Strip leading hyphen.
                        if ( asProfilePromptKey.StartsWith("-") )
                            asProfilePromptKey = asProfilePromptKey.Substring(1, asProfilePromptKey.Length - 1);

                        // Insert prefix.
                        asProfilePromptKey = msProfilePromptKeyPrefix + asProfilePromptKey;
                    }

                    // Make the answer key from the prompt key and the prompt key suffix.
                    lsPromptAnswerKey = asProfilePromptKey + msProfilePromptKeySuffix;

                    // Only the first display of a modeless dialog can contain a checkbox.
                    // Why? Because the first prompt is not modeless. That's the only way
                    // to capture the checkbox value. BTW, "lbUseCheckBox" is reset here
                    // for use outside of this block to avoid the default setting next.
                    if ( abShowModeless )
                        lbUseCheckBox = !aoProfile.ContainsKey(asProfilePromptKey);

                    if (      !aoProfile.bValue(asProfilePromptKey, false)
                            && aoProfile.ContainsKey(lsPromptAnswerKey) )
                    {
                        // Do not prompt. Return the previous stored answer instead.
                        return (tvMessageBoxResults)aoProfile.iValue(
                                lsPromptAnswerKey, (int)tvMessageBoxResults.None);
                    }
                }

        if ( null == asMessageCaption )
        {
            // No caption provided. Let's try to get one another way.

            if ( null != aoWindow )             // Try window title first.
                asMessageCaption = aoWindow.Title;
            else
            if ( null != Application.Current && null != Application.Current.MainWindow )  // Next try for application name.
                asMessageCaption = Application.Current.MainWindow.Name;
            else
                asMessageCaption = System.IO.Path.GetFileNameWithoutExtension(Application.ResourceAssembly.Location);
        }

        if ( null != aoWindow )
            aoWindow.Cursor = null;             // Turn off wait cursor in parent window.

        tvMessageBox    loMsgBox = new tvMessageBox();
                        loMsgBox.MessageText.Text = asMessageText;

                        // Use some parent window attributes, if available.
                        if ( null != aoWindow )
                        {
                            // Use the parent window's icon.
                            loMsgBox.Icon = aoWindow.Icon;

                            // Use the given asMessageCaption as the MsgBox title, if not null.
                            // Otherwise use the parent window title with an added question mark.
                            loMsgBox.Title = null != asMessageCaption
                                    ? asMessageCaption : aoWindow.Title + "?";
                        }

                        // Display the MsgBox header / title (ie. the caption), if provided.
                        if ( null != asMessageCaption )
                        {
                            loMsgBox.MessageTitle.Content = asMessageCaption;
                            loMsgBox.MessageTitle.Visibility = Visibility.Visible;
                        }

                        loMsgBox.SelectButtons(aeTvMessageBoxButtons);
                        loMsgBox.SelectIcon(aeTvMessageBoxIcon);

                        if ( lbUseCheckBox )
                        {
                            switch (aeTvMessageBoxCheckBoxType)
                            {
                                case tvMessageBoxCheckBoxTypes.DontAsk:
                                    loMsgBox.chkDontAsk.Visibility = Visibility.Visible;
                                    break;
                                case tvMessageBoxCheckBoxTypes.SkipThis:
                                    loMsgBox.chkSkipThis.Visibility = Visibility.Visible;
                                    break;
                            }
                        }

                        if ( !abShowModeless )
                        {
                            loMsgBox.ShowDialog();
                        }
                        else
                        {
                            // It can only be modeless after the checkbox has been stored.
                            if ( lbUseCheckBox )
                                loMsgBox.ShowDialog();
                            else
                                loMsgBox.Show();
                        }

                        if ( lbUseCheckBox )
                        {
                            bool lbCheckBoxValue = false;

                            switch (aeTvMessageBoxCheckBoxType)
                            {
                                case tvMessageBoxCheckBoxTypes.DontAsk:
                                    lbCheckBoxValue = (bool)loMsgBox.chkDontAsk.IsChecked;
                                    break;
                                case tvMessageBoxCheckBoxTypes.SkipThis:
                                    lbCheckBoxValue = (bool)loMsgBox.chkSkipThis.IsChecked;
                                    break;
                            }

                            // Use the answer override whenever not "none". This value is
                            // necessary for certain stored answers that don't make sense
                            // in a given context (eg. both "skip this" and "cancel" selected).
                            if ( tvMessageBoxResults.None == aeTvMessageBoxResultsOverride )
                                aeTvMessageBoxResultsOverride = loMsgBox.eTvMessageBoxResult;

                            // Reverse the boolean. "Don't ask" or "Skip this" means "Don't prompt".
                            aoProfile[asProfilePromptKey] = !lbCheckBoxValue;
                            aoProfile[lsPromptAnswerKey] = (int)aeTvMessageBoxResultsOverride;
                            aoProfile.Save();
                        }

        liTvMessageBoxResult = loMsgBox.eTvMessageBoxResult;

        return liTvMessageBoxResult;
    }

    public static void ShowBriefly(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvMessageBoxIcons aeTvMessageBoxIcon
            , int aiCloseAfterMS
            )
    {
        tvMessageBox    loMsgBox = new tvMessageBox();
                        loMsgBox.MessageText.Text = asMessageText;

                        // Use some parent window attributes, if available.
                        if ( null != aoWindow )
                        {
                            // Use the parent window's icon.
                            loMsgBox.Icon = aoWindow.Icon;

                            // Use the given asMessageCaption as the MsgBox title, if not null.
                            // Otherwise use the parent window title with an added question mark.
                            loMsgBox.Title = null != asMessageCaption
                                    ? asMessageCaption : aoWindow.Title + "?";
                        }

                        // Display the MsgBox header / title (ie. the caption), if provided.
                        if ( null != asMessageCaption )
                        {
                            loMsgBox.MessageTitle.Content = asMessageCaption;
                            loMsgBox.MessageTitle.Visibility = Visibility.Visible;
                        }

                        loMsgBox.SelectButtons(tvMessageBoxButtons.OK);
                        loMsgBox.SelectIcon(aeTvMessageBoxIcon);
                        loMsgBox.Show();

        if ( aiCloseAfterMS > 0 )
        {
            DateTime ldtCloseTime = DateTime.Now.AddMilliseconds(aiCloseAfterMS);

            while ( DateTime.Now < ldtCloseTime )
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(200);

                if ( loMsgBox.bDialogAccepted )
                    break;
            }

            loMsgBox.Close();
        }
    }

    public void ShowWait(Window aoWindow, string asMessageText, int aiStayOpenAtLeastMS)
    {
        this.Cursor = Cursors.Wait;
        mbIsShowWait = true;
        this.BottomPanel.Height = 35;
        this.MessageText.Text = asMessageText;

        // Use some parent window attributes, if available.
        if ( null != aoWindow )
        {
            // Use the parent window's icon.
            this.Icon = aoWindow.Icon;

            // Use the parent window title with an added message.
            this.Title = aoWindow.Title + " - wait ...";
        }

        this.Show();

        if ( aiStayOpenAtLeastMS > 0 )
        {
            DateTime ldtOpenTime = DateTime.Now.AddMilliseconds(aiStayOpenAtLeastMS);

            while ( DateTime.Now < ldtOpenTime )
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(200);
            }
        }
    }

    public static void ShowError(Window aoWindow, Exception aoException)
    {
        tvMessageBox.Show(
                  aoWindow
                , aoException.Message
                , null
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Error
                );
    }

    public static void ShowError(Window aoWindow, string asMessageText)
    {
        tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , null
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Error
                );
    }

    public static void ShowError(Window aoWindow, string asMessageText, string asMessageCaption)
    {
        tvMessageBox.Show(aoWindow, asMessageText, asMessageCaption
                , tvMessageBoxButtons.OK, tvMessageBoxIcons.Error);
    }

    public static void ShowModelessError(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvProfile aoProfile
            )
    {
        tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Error
                , true
                );
    }

    public static void ShowModelessError(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvProfile aoProfile
            , string asProfilePromptKey
            )
    {
        tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , asMessageCaption
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Error
                , true
                , tvMessageBoxCheckBoxTypes.SkipThis
                , aoProfile
                , asProfilePromptKey
                , tvMessageBoxResults.None
                );
    }

    public static void ShowWarning(Window aoWindow, string asMessageText)
    {
        tvMessageBox.Show(
                  aoWindow
                , asMessageText
                , null
                , tvMessageBoxButtons.OK
                , tvMessageBoxIcons.Warning
                );
    }

    public static void ShowWarning(Window aoWindow, string asMessageText, string asMessageCaption)
    {
        tvMessageBox.Show(aoWindow, asMessageText, asMessageCaption
                , tvMessageBoxButtons.OK, tvMessageBoxIcons.Warning);
    }

    public static void ResetAllPrompts(
              Window aoWindow
            , string asMessageText
            , string asMessageCaption
            , tvProfile aoProfile
            )
    {
        if ( tvMessageBoxResults.Yes == tvMessageBox.Show(
                      aoWindow
                    , asMessageText
                    , asMessageCaption
                    , tvMessageBoxButtons.YesNo
                    , tvMessageBoxIcons.Alert
                    )
                )
        {
            aoProfile.Remove(msProfilePromptKeyPrefix + "*");
            aoProfile.Save();
        }
    }
}
