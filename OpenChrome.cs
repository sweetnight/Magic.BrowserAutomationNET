namespace Magic.BrowserAutomationNET
{
    public sealed class OpenChrome : IDisposable
    {

        public enum StateCode
        {
            Opened,
            NotOpened
        } // end of enum

        public enum EventType
        {
            Start,
            ChromeIsConnected,
            NewChromeIsOpened,
            BrowserClosed,
            ZombieKilled
        } // end of enum

        public event Action<OpenChromeEventArgs>? OpenChromeEvents;

        public class OpenChromeEventArgs
        {
            public EventType EventType { get; set; }
            public Chrome Chrome { get; set; }

            public OpenChromeEventArgs(EventType eventType, Chrome chrome)
            {

                EventType = eventType;
                Chrome = chrome;

            } // end of method
        } // end of method

        private bool _disposed = false; // untuk mendeteksi double dispose

        public Chrome Chrome { get; }
        public CancellationToken CancellationToken { get; }

        public StateCode ChromeInitialState { get; private set; }

        public OpenChrome(Chrome chrome, CancellationToken cancellationToken)
        {

            Chrome = chrome;
            CancellationToken = cancellationToken;

        } // end of method

        public Chrome Start()
        {

            OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.Start, Chrome));

            if (Chrome!.Driver == null)
            {
                ChromeInitialState = StateCode.NotOpened;
                return ChromeIsNotConnected();
            }

            ChromeInitialState = StateCode.Opened;

            try
            {
                string title = Chrome.Driver.Title;

                OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.ChromeIsConnected, Chrome));

                return Chrome;
            }
            catch
            {
                return ChromeIsNotConnected();
            }

        } // end of method

        private Chrome ChromeIsNotConnected()
        {

            // CLOSE DULU CHROME TERKAIT YANG SUDAH TERBUKA TAPI TIDAK CONNECT

            bool killedOpenedChrome = Chrome!.KillBrowserByUserDataDir(Chrome.UserDataDir);
            OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.ZombieKilled, Chrome));

            CancellationToken.ThrowIfCancellationRequested();

            Chrome.OpenBrowser();
            OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.NewChromeIsOpened, Chrome));

            return Chrome;

        } // end of method

        public void ForceClose()
        {
            if (Chrome != null)
            {
                Chrome.CloseBrowser();
                OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.BrowserClosed, Chrome));
            }

            Dispose();
        } // end of method

        public bool DisposeIfOwned()
        {

            if (ChromeInitialState == StateCode.NotOpened)
            {
                Chrome.CloseBrowser();
                OpenChromeEvents?.Invoke(new OpenChromeEventArgs(EventType.BrowserClosed, Chrome));

                Dispose();

                return true;
            }

            return false;

        } // end of method

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            // cleanup jika ada
            GC.SuppressFinalize(this);
        }

    } // end of class
} // end of namespace
