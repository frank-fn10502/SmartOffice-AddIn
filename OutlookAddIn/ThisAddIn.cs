using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using OutlookAddIn.Clients;
using OutlookAddIn.Infrastructure.Threading;
using OutlookAddIn.Ribbon;
using OutlookAddIn.UI;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _chatTaskPane;
        private ChatPane _chatPane;
        private SmartOfficeRibbon _ribbon;
        private SignalRClient _signalRClient;
        private OutlookThreadInvoker _outlookThread;
        private readonly SemaphoreSlim _commandGate = new SemaphoreSlim(1, 1);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _chatPane = new ChatPane();
                _chatTaskPane = this.CustomTaskPanes.Add(_chatPane, "SmartOffice Chat");
                _chatTaskPane.Width = 320;
                _chatTaskPane.Visible = false;
                _outlookThread = new OutlookThreadInvoker(_chatPane);

                // Defer SignalR connection to avoid blocking Outlook startup
                var startupTimer = new Timer();
                startupTimer.Interval = 1500; // let Outlook finish loading first
                startupTimer.Tick += (s, ev) =>
                {
                    startupTimer.Stop();
                    startupTimer.Dispose();
                    InitSignalRAsync();
                };
                startupTimer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SmartOffice startup error: " + ex);
            }
        }

        private async void InitSignalRAsync()
        {
            try
            {
                _signalRClient = new SignalRClient(HubClient.BaseUrl);
                _signalRClient.CommandReceived += OnCommandReceivedAsync;
                await _signalRClient.StartAsync();
                _chatPane.SetSignalRClient(_signalRClient);
                System.Diagnostics.Debug.WriteLine("SmartOffice SignalR connected to /hub/outlook-addin.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SmartOffice SignalR connection failed: " + ex.Message);
            }
        }

        public void ToggleChatPane()
        {
            if (_chatTaskPane != null)
                _chatTaskPane.Visible = !_chatTaskPane.Visible;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SmartOfficeRibbon();
            return _ribbon;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_signalRClient != null)
            {
                _signalRClient.Dispose();
                _signalRClient = null;
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
