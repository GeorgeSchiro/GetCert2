//-----------------------------------------------------------------------------------------
// Author:   Murray Foxcroft - April 2009
//-----------------------------------------------------------------------------------------
using System;
using System.Windows.Forms;

namespace GetCert2
{
    /// <summary>
    /// An extension of the notifyIcon Windows Forms class, unfortunately its a 
    //  sealed class so it cannot be inherited. This class adds a timer and additional 
    //  methods and events to allow for monitoring when a mouse enters and leaves the icon area. 
    /// </summary>
    /// 
    public class ExtendedNotifyIcon : IDisposable
    {
        public NotifyIcon targetNotifyIcon;
        private System.Drawing.Point notifyIconMousePosition;
        private Timer delayMouseLeaveEventTimer;
        private bool bStopMouseMoveEventFromFiring;
    
        public delegate void MouseLeaveHandler();
        public event MouseLeaveHandler MouseLeave;

        public delegate void MouseMoveHandler();
        public event MouseMoveHandler MouseMove;

        public delegate void MouseClickHandler();
        public event MouseClickHandler MouseClick;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="millisecondsToDelayMouseLeaveEvent"></param>
        public ExtendedNotifyIcon(int millisecondsToDelayMouseLeaveEvent)
        {
            // Configure and show a notification icon in the system tray
            targetNotifyIcon = new NotifyIcon();
            targetNotifyIcon.Visible = true;
            targetNotifyIcon.MouseMove += new MouseEventHandler(targetNotifyIcon_MouseMove);
            targetNotifyIcon.MouseClick += new MouseEventHandler(targetNotifyIcon_MouseClick);

            delayMouseLeaveEventTimer = new Timer();
            delayMouseLeaveEventTimer.Tick += new EventHandler(delayMouseLeaveEventTimer_Tick);
            delayMouseLeaveEventTimer.Interval = 100;
        }

        /// <summary>
        /// Chained constructor - default millisecondsToDelayMouseLeaveEvent is 100ms
        /// </summary>
        public ExtendedNotifyIcon() : this(100) { }

        /// <summary>
        /// Manual override exposed - START the timer which will ultimately trigger the mouse leave event
        /// </summary>
        public void StartMouseLeaveTimer()
        {
            delayMouseLeaveEventTimer.Start();
        }

        /// <summary>
        /// Manual override exposed - STOP the timer that would ultimately close the window
        /// </summary>
        public void StopMouseLeaveEventFromFiring()
        {
            delayMouseLeaveEventTimer.Stop();
        }

        /// <summary>
        /// Manual override exposed - disable move events
        /// </summary>
        public void StopMouseMoveEventFromFiring()
        {
            bStopMouseMoveEventFromFiring = true;
        }

        /// <summary>
        /// If the mouse is moving over the notify icon, the popup must be displayed.     
        /// Note: There is no event on the notify icon to trap when the mouse leave, so a timer is used in conjunction 
        /// with tracking the position of the mouse to test for when the popup window needs to be closed. See timer tick event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void targetNotifyIcon_MouseMove(object sender, MouseEventArgs e)
        {
            if ( !bStopMouseMoveEventFromFiring )
            {
                notifyIconMousePosition = System.Windows.Forms.Control.MousePosition; // Track the position of the mouse over the notify icon
                MouseMove(); // The mouse is moving over the notify Icon, raise the event
                delayMouseLeaveEventTimer.Start();  // The timer counts down and closes the window, as the mouse moves over the icon, keep starting (resetting) this to stop it from closing the popup
            }
        }

        public void targetNotifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            notifyIconMousePosition = System.Windows.Forms.Control.MousePosition; // Track the position of the mouse over the notify icon
            MouseClick(); // The mouse is clicked on the notify Icon, raise the event
        }

        /// <summary>
        /// Under the right conditions, raise the event to the popup window to tell it to close.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void delayMouseLeaveEventTimer_Tick(object sender, EventArgs e)
        {
            // If the mouse position over the icon does not match the sryceen position, the mouse has left the icon (think of this as a type of hit test) 
            if (notifyIconMousePosition != System.Windows.Forms.Control.MousePosition)
            {
                MouseLeave();  // Raise the event for the mouse leaving 
                delayMouseLeaveEventTimer.Stop(); // Stop the timer, no longer reqired.
            }
        }

        #region IDisposable Members

        /// <summary>
        /// Standard IDisposable interface implementation. If you dont dispose the windows notify icon, the application
        /// closes but the icon remains in the task bar until such time as you mouse over it.
        /// </summary>
        private bool _IsDisposed = false;

        ~ExtendedNotifyIcon()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            // Tell the garbage collector not to call the finalizer
            // since all the cleanup will already be done.
            GC.SuppressFinalize(true);
        }

        protected virtual void Dispose(bool IsDisposing)
        {
            if (_IsDisposed)
                return;

            if (IsDisposing)
            {
                targetNotifyIcon.Dispose();
            }

            // Free any unmanaged resources in this section
            _IsDisposed = true;

        #endregion
        }
    }
}
