using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAddIn.Clients;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.UI
{
    public partial class ChatPane : UserControl
    {
        private SignalRClient _signalRClient;

        public ChatPane()
        {
            InitializeComponent();
            this.txtChat.KeyDown += TxtChat_KeyDown;
        }

        /// <summary>
        /// Set the SignalR client reference so chat uses SendChatMessage via SignalR.
        /// </summary>
        public void SetSignalRClient(SignalRClient client)
        {
            _signalRClient = client;
        }

        private void TxtChat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnSend_Click(sender, e);
            }
        }

        private async void btnSend_Click(object sender, EventArgs e)
        {
            var text = txtChat.Text.Trim();
            if (string.IsNullOrEmpty(text)) return;
            txtChat.Clear();

            try
            {
                if (_signalRClient != null && _signalRClient.IsConnected)
                {
                    await _signalRClient.SendChatMessageAsync(new ChatMessageDto
                    {
                        Id = Guid.NewGuid().ToString(),
                        Source = "outlook",
                        Text = text,
                        Timestamp = DateTime.Now
                    });
                }
                else
                {
                    AppendMessage("Error", "SignalR not connected");
                }
            }
            catch (Exception ex)
            {
                AppendMessage("Error", ex.Message);
            }
        }

        public void AppendMessage(string source, string text)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => AppendMessage(source, text)));
                return;
            }
            rtbHistory.AppendText("[" + source + "] " + text + "\n");
            rtbHistory.ScrollToCaret();
        }
    }
}
