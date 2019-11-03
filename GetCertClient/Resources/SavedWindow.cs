using System;
using System.ComponentModel;
using System.Windows;

namespace GetCert2
{
    public class SavedWindow : Window
    {
        private bool mbLoading;


        public DoGetCert oApp
        {
            get
            {
                return (DoGetCert)Application.Current;
            }
        }

        public string TN
        {
            get
            {
                return this.GetType().Name;
            }
        }


        private void Window_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                // Update the window settings.
                if ( WindowState.Normal == this.WindowState )
                {
                    this.oApp.oProfile[string.Format("-{0}.Top", this.TN)] = this.Top;
                    this.oApp.oProfile[string.Format("-{0}.Left", this.TN)] = this.Left;
                    this.oApp.oProfile.Save();
                }

                this.SaveWindowSettings();
            }
            catch (Exception ex)
            {
                tvMessageBox.ShowError(this, ex);
            }
        }

        private void LoadWindowSettings()
        {
            mbLoading = true;

            try
            {
                string lsWindowStateKey = string.Format("-{0}.WindowState", this.TN);

                if ( this.oApp.oProfile.ContainsKey(lsWindowStateKey) )
                {
                    this.WindowStartupLocation = WindowStartupLocation.Manual;
                    this.WindowState = (WindowState) this.oApp.oProfile.iValue(lsWindowStateKey, 0);
                    this.Top = this.oApp.oProfile.dValue(string.Format("-{0}.Top", this.TN), 100);
                    this.Left = this.oApp.oProfile.dValue(string.Format("-{0}.Left", this.TN), 100);
                }
            }
            catch (Exception ex)
            {
                tvMessageBox.ShowError(this, ex);
            }
            finally
            {
                mbLoading = false;
            }
        }

        private void SaveWindowSettings()
        {
            if ( mbLoading )
                return;

            this.oApp.oProfile.Save();
        }


        public void Init()
        {
            this.Closing += new CancelEventHandler(this.Window_Closing);

            try
            {
                this.LoadWindowSettings();
            }
            catch (Exception ex)
            {
                tvMessageBox.ShowError(this, ex);
            }
        }
    }
}
