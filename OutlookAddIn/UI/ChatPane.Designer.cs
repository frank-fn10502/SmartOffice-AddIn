namespace OutlookAddIn.UI
{
    partial class ChatPane
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtChat;
        private System.Windows.Forms.RichTextBox rtbHistory;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Panel panelInput;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.rtbHistory = new System.Windows.Forms.RichTextBox();
            this.panelInput = new System.Windows.Forms.Panel();
            this.txtChat = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.panelInput.SuspendLayout();
            this.SuspendLayout();

            // rtbHistory
            this.rtbHistory.BackColor = System.Drawing.Color.White;
            this.rtbHistory.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rtbHistory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbHistory.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rtbHistory.ReadOnly = true;
            this.rtbHistory.Location = new System.Drawing.Point(0, 0);
            this.rtbHistory.Name = "rtbHistory";
            this.rtbHistory.Size = new System.Drawing.Size(300, 400);

            // panelInput
            this.panelInput.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelInput.Height = 36;
            this.panelInput.Controls.Add(this.txtChat);
            this.panelInput.Controls.Add(this.btnSend);
            this.panelInput.Padding = new System.Windows.Forms.Padding(2);

            // txtChat
            this.txtChat.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtChat.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtChat.Name = "txtChat";

            // btnSend
            this.btnSend.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnSend.Text = "Send";
            this.btnSend.Width = 60;
            this.btnSend.Name = "btnSend";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);

            // ChatPane
            this.Controls.Add(this.rtbHistory);
            this.Controls.Add(this.panelInput);
            this.Name = "ChatPane";
            this.Size = new System.Drawing.Size(300, 436);
            this.panelInput.ResumeLayout(false);
            this.panelInput.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
