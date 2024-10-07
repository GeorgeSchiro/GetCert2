using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

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
            this.SaveWindowSettings();
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            this.SaveWindowSettings();
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
                    this.WindowState = (WindowState)Enum.Parse(typeof(WindowState), this.oApp.oProfile.sValue(lsWindowStateKey, "Normal"), true);
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

            try
            {
                // Update window settings.
                if ( WindowState.Normal == this.WindowState )
                {
                    this.oApp.oProfile[string.Format("-{0}.WindowState", this.TN)] = this.WindowState.ToString();
                    this.oApp.oProfile[string.Format("-{0}.Top", this.TN)] = this.Top;
                    this.oApp.oProfile[string.Format("-{0}.Left", this.TN)] = this.Left;
                    this.oApp.oProfile.Save();
                }

                this.oApp.oProfile.Save();
            }
            catch (Exception ex)
            {
                tvMessageBox.ShowError(this, ex);
            }
        }


        public void Init()
        {
            this.Closing += new CancelEventHandler(this.Window_Closing);
            this.MouseLeftButtonUp += new MouseButtonEventHandler(this.Window_MouseLeftButtonUp);

            this.LoadWindowSettings();
        }
    }
}
