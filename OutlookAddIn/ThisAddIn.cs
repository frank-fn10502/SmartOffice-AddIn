using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _chatTaskPane;
        private ChatPane _chatPane;
        private SmartOfficeRibbon _ribbon;
        private SignalRClient _signalRClient;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _chatPane = new ChatPane();
                _chatTaskPane = this.CustomTaskPanes.Add(_chatPane, "SmartOffice Chat");
                _chatTaskPane.Width = 320;
                _chatTaskPane.Visible = false;

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

        /// <summary>
        /// Dispatches incoming OutlookCommand from Hub.
        /// Heavy Outlook COM work is marshalled back to the UI thread via BeginInvoke
        /// to avoid freezing the STA thread.
        /// </summary>
        private async Task OnCommandReceivedAsync(OutlookCommand cmd)
        {
            // Marshal to UI thread for COM access
            var tcs = new TaskCompletionSource<bool>();
            _chatPane.BeginInvoke((Action)(async () =>
            {
                try
                {
                    await HandleCommandAsync(cmd);
                }
                catch (Exception ex)
                {
                    try { await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "Error: " + SanitizeExceptionForLog(ex)); } catch { }
                }
                finally
                {
                    tcs.TrySetResult(true);
                }
            }));
            await tcs.Task;
        }

        private async Task HandleCommandAsync(OutlookCommand cmd)
        {
            try { await _signalRClient.ReportLogAsync("info", "Received command: " + cmd.Type); } catch { }

            switch (cmd.Type)
            {
                case "ping":
                    // Only report success when Outlook object model is actually callable
                    bool outlookReady = false;
                    try
                    {
                        var ns = this.Application.Session;
                        outlookReady = ns != null && ns.Stores != null;
                    }
                    catch { }
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, outlookReady, outlookReady ? "pong" : "Outlook session not ready");
                    break;

                case "fetch_folder_roots":
                    await HandleFetchFolderRootsAsync(cmd);
                    break;

                case "fetch_folder_children":
                    await HandleFetchFolderChildrenAsync(cmd);
                    break;

                case "fetch_mails":
                    try
                    {
                        var mr = cmd.MailsRequest;
                        var mailReq = new FetchMailsRequest
                        {
                            FolderPath = mr?.FolderPath ?? "",
                            Range = mr?.Range ?? "30d",
                            MaxCount = (mr?.MaxCount > 0 ? mr.MaxCount : 100),
                            ReceivedFrom = mr?.ReceivedFrom,
                            ReceivedTo = mr?.ReceivedTo
                        };
                        if (mailReq.MaxCount > 500) mailReq.MaxCount = 500;

                        List<MailItemDto> mails = ReadMails(mailReq);

                        await _signalRClient.PushMailsAsync(mails ?? new List<MailItemDto>());
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mails completed. Items: " + (mails?.Count ?? 0));
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mails error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "fetch_mail_search_slice":
                    await HandleMailSearchSliceAsync(cmd);
                    break;

                case "fetch_folder_mails_slice":
                    await HandleFolderMailsSliceAsync(cmd);
                    break;

                case "fetch_mail_body":
                    try
                    {
                        var bodyReq = cmd.MailBodyRequest;
                        if (bodyReq == null || string.IsNullOrEmpty(bodyReq.MailId))
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_body failed: missing mail id");
                            break;
                        }

                        MailBodyDto bodyDto = ReadMailBody(bodyReq.MailId, bodyReq.FolderPath);

                        if (bodyDto == null)
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_body failed: mail not found");
                            break;
                        }

                        await _signalRClient.PushMailBodyAsync(bodyDto);
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mail_body completed.");
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_body error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "fetch_mail_attachments":
                    try
                    {
                        var attReq = cmd.MailAttachmentsRequest;
                        if (attReq == null || string.IsNullOrEmpty(attReq.MailId))
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_attachments failed: missing mail id");
                            break;
                        }

                        MailAttachmentsDto attDto = ReadMailAttachments(attReq.MailId, attReq.FolderPath);

                        if (attDto == null)
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_attachments failed: mail not found");
                            break;
                        }

                        await _signalRClient.PushMailAttachmentsAsync(attDto);
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mail_attachments completed. Items: " + (attDto.Attachments?.Count ?? 0));
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_attachments error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "export_mail_attachment":
                    try
                    {
                        var expReq = cmd.ExportMailAttachmentRequest;
                        if (expReq == null || string.IsNullOrEmpty(expReq.MailId))
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "export_mail_attachment failed: missing mail id");
                            break;
                        }

                        ExportedMailAttachmentDto expDto = ExportMailAttachment(expReq);

                        if (expDto == null)
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "export_mail_attachment failed: attachment not found or export error");
                            break;
                        }

                        await _signalRClient.PushExportedMailAttachmentAsync(expDto);
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "export_mail_attachment completed.");
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "export_mail_attachment error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "fetch_rules":
                    try
                    {
                        await _signalRClient.ReportLogAsync("info", "fetch_rules: starting");
                        List<OutlookRuleDto> rules = ReadRules();
                        await _signalRClient.PushRulesAsync(rules ?? new List<OutlookRuleDto>());
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_rules completed. Items: " + (rules?.Count ?? 0));
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_rules error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "manage_rule":
                    await HandleManageRuleAsync(cmd);
                    break;

                case "fetch_categories":
                    try
                    {
                        await _signalRClient.ReportLogAsync("info", "fetch_categories: starting");
                        List<OutlookCategoryDto> categories = ReadCategories();
                        await _signalRClient.PushCategoriesAsync(categories ?? new List<OutlookCategoryDto>());
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_categories completed. Items: " + (categories?.Count ?? 0));
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_categories error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "fetch_calendar":
                    try
                    {
                        DateTime calStart, calEnd;
                        if (cmd.CalendarRequest != null && !string.IsNullOrEmpty(cmd.CalendarRequest.StartDate) && !string.IsNullOrEmpty(cmd.CalendarRequest.EndDate))
                        {
                            if (!DateTime.TryParse(cmd.CalendarRequest.StartDate, out calStart))
                                calStart = DateTime.Now.Date;
                            if (!DateTime.TryParse(cmd.CalendarRequest.EndDate, out calEnd))
                                calEnd = calStart.AddMonths(1);
                        }
                        else
                        {
                            int days = cmd.CalendarRequest?.DaysForward ?? 14;
                            if (days <= 0) days = 14;
                            calStart = DateTime.Now.Date;
                            calEnd = calStart.AddDays(days);
                        }

                        await _signalRClient.ReportLogAsync("info", $"fetch_calendar: {calStart:yyyy-MM-dd} to {calEnd:yyyy-MM-dd}");

                        List<CalendarEventDto> events = ReadCalendarEvents(calStart, calEnd);

                        await _signalRClient.PushCalendarAsync(events ?? new List<CalendarEventDto>());
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_calendar completed. Items: " + (events?.Count ?? 0));
                    }
                    catch (Exception ex)
                    {
                        await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_calendar error: " + SanitizeExceptionForLog(ex));
                    }
                    break;

                case "update_mail_properties":
                    await HandleUpdateMailPropertiesAsync(cmd);
                    break;

                case "move_mail":
                    await HandleMoveMailAsync(cmd);
                    break;

                case "move_mails":
                    await HandleMoveMailsAsync(cmd);
                    break;

                case "delete_mail":
                    await HandleDeleteMailAsync(cmd);
                    break;

                case "create_folder":
                    await HandleCreateFolderAsync(cmd);
                    break;

                case "delete_folder":
                    await HandleDeleteFolderAsync(cmd);
                    break;

                case "upsert_category":
                    await HandleUpsertCategoryAsync(cmd);
                    break;

                default:
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "Unknown command type: " + cmd.Type);
                    break;
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
