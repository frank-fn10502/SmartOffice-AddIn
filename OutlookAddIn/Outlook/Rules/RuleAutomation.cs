using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using OutlookAddIn.Clients;
using OutlookAddIn.Infrastructure.Diagnostics;
using OutlookAddIn.Infrastructure.Threading;
using OutlookAddIn.OutlookServices.Folders;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Rules
{
    internal sealed class OutlookRuleCommandHandler
    {
        private readonly SignalRClient _signalRClient;
        private readonly OutlookThreadInvoker _outlookThread;
        private readonly Outlook.Application _application;
        private readonly OutlookRuleReader _ruleReader;
        private readonly OutlookFolderLocator _folderLocator;

        public OutlookRuleCommandHandler(
            SignalRClient signalRClient,
            OutlookThreadInvoker outlookThread,
            Outlook.Application application)
        {
            _signalRClient = signalRClient ?? throw new ArgumentNullException(nameof(signalRClient));
            _outlookThread = outlookThread ?? throw new ArgumentNullException(nameof(outlookThread));
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _ruleReader = new OutlookRuleReader(application);
            _folderLocator = new OutlookFolderLocator(application);
        }

        public async Task HandleManageRuleAsync(OutlookCommand cmd)
        {
            var req = cmd.RuleRequest;
            if (req == null || string.IsNullOrEmpty(req.Operation))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "manage_rule failed: operation is required").ConfigureAwait(false);
                return;
            }

            List<OutlookRuleDto> updatedRules;
            try
            {
                updatedRules = await _outlookThread.InvokeAsync(() =>
                {
                    ApplyManageRule(req);
                    return _ruleReader.ReadRules();
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "manage_rule error: " + SensitiveLogSanitizer.Sanitize(ex)).ConfigureAwait(false);
                return;
            }

            try
            {
                await _signalRClient.PushRulesAsync(updatedRules ?? new List<OutlookRuleDto>()).ConfigureAwait(false);
            }
            catch { }

            await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                $"manage_rule {req.Operation} completed").ConfigureAwait(false);
        }

        private void ApplyManageRule(OutlookCommandRuleRequest req)
        {
            Outlook.Store store = null;
            Outlook.Rules rules = null;
            try
            {
                if (!string.IsNullOrEmpty(req.StoreId))
                {
                    Outlook.Stores stores = null;
                    try
                    {
                        stores = _application.Session.Stores;
                        foreach (Outlook.Store candidate in stores)
                        {
                            string storeId = "";
                            try { storeId = candidate.StoreID ?? ""; } catch { }
                            if (string.Equals(storeId, req.StoreId, StringComparison.OrdinalIgnoreCase))
                            {
                                store = candidate;
                                break;
                            }

                            ReleaseComObjectQuietly(candidate);
                        }
                    }
                    finally
                    {
                        ReleaseComObjectQuietly(stores);
                    }

                    if (store == null)
                        throw new InvalidOperationException("manage_rule: requested Outlook store was not found");
                }

                if (store == null)
                    store = _application.Session.DefaultStore;

                if (store == null)
                    throw new InvalidOperationException("Cannot resolve Outlook store for manage_rule");

                rules = store.GetRules();
                if (rules == null)
                    throw new InvalidOperationException("GetRules returned null");

                string operation = req.Operation?.ToLower() ?? "";
                switch (operation)
                {
                    case "upsert":
                        ApplyRuleUpsert(rules, req);
                        break;
                    case "delete":
                        ApplyRuleDelete(rules, req);
                        break;
                    case "set_enabled":
                        ApplyRuleSetEnabled(rules, req);
                        break;
                    default:
                        throw new NotSupportedException($"manage_rule: unsupported operation '{req.Operation}'");
                }

                rules.Save(false);
            }
            finally
            {
                ReleaseComObjectQuietly(rules);
                ReleaseComObjectQuietly(store);
            }
        }

        private void ApplyRuleUpsert(Outlook.Rules rules, OutlookCommandRuleRequest req)
        {
            Outlook.Rule existingRule = null;
            if (req.OriginalExecutionOrder.HasValue && req.OriginalExecutionOrder.Value > 0)
            {
                try
                {
                    var rule = rules[req.OriginalExecutionOrder.Value];
                    if (rule != null)
                        existingRule = rule;
                }
                catch { }
            }

            if (existingRule == null && !string.IsNullOrEmpty(req.OriginalRuleName))
            {
                try
                {
                    for (int i = 1; i <= rules.Count; i++)
                    {
                        Outlook.Rule rule = null;
                        try
                        {
                            rule = rules[i];
                            if (string.Equals(rule.Name, req.OriginalRuleName, StringComparison.OrdinalIgnoreCase))
                            {
                                existingRule = rule;
                                break;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (rule != null && !ReferenceEquals(rule, existingRule))
                                ReleaseComObjectQuietly(rule);
                        }
                    }
                }
                catch { }
            }

            Outlook.Rule targetRule = existingRule;
            try
            {
                if (targetRule == null)
                {
                    var ruleName = (req.RuleName ?? "").Trim();
                    if (string.IsNullOrEmpty(ruleName))
                        throw new InvalidOperationException("manage_rule upsert: rule name is required");

                    var ruleType = string.Equals(req.RuleType, "send", StringComparison.OrdinalIgnoreCase)
                        ? Outlook.OlRuleType.olRuleSend
                        : Outlook.OlRuleType.olRuleReceive;
                    targetRule = rules.Create(ruleName, ruleType);
                }
                else if (!string.IsNullOrEmpty(req.RuleName))
                {
                    targetRule.Name = req.RuleName;
                }

                targetRule.Enabled = req.Enabled;
                if (req.ExecutionOrder.HasValue && req.ExecutionOrder.Value > 0)
                    targetRule.ExecutionOrder = req.ExecutionOrder.Value;

                if (req.Conditions != null)
                {
                    Outlook.RuleConditions conditions = null;
                    try
                    {
                        conditions = targetRule.Conditions;
                        ApplyTextRuleCondition(conditions.Subject, req.Conditions.SubjectContains, existingRule != null, "subject");
                        ApplyTextRuleCondition(conditions.Body, req.Conditions.BodyContains, existingRule != null, "body");
                        ApplySenderAddressCondition(conditions.SenderAddress, req.Conditions.SenderAddressContains, existingRule != null);
                        ApplyCategoryCondition(conditions.Category, req.Conditions.Categories, existingRule != null);
                        ApplyHasAttachmentCondition(conditions.HasAttachment, req.Conditions.HasAttachment, existingRule != null);
                    }
                    finally
                    {
                        ReleaseComObjectQuietly(conditions);
                    }
                }

                if (req.Actions != null)
                {
                    Outlook.RuleActions actions = null;
                    try
                    {
                        actions = targetRule.Actions;
                        ApplyMoveToFolderAction(actions.MoveToFolder, req.Actions.MoveToFolderPath, existingRule != null);
                        ApplyAssignToCategoryAction(actions.AssignToCategory, req.Actions.AssignCategories, existingRule != null);
                        ApplyMarkAsTaskAction(actions.MarkAsTask, req.Actions.MarkAsTask, existingRule != null);
                        ApplyStopAction(actions.Stop, req.Actions.StopProcessingMoreRules);
                    }
                    finally
                    {
                        ReleaseComObjectQuietly(actions);
                    }
                }
            }
            finally
            {
                ReleaseComObjectQuietly(targetRule);
            }
        }

        private void ApplyTextRuleCondition(object condition, List<string> values, bool disableWhenEmpty, string label)
        {
            ApplyStringArrayRulePart(condition, "Text", values, disableWhenEmpty, label);
        }

        private void ApplySenderAddressCondition(object condition, List<string> values, bool disableWhenEmpty)
        {
            ApplyStringArrayRulePart(condition, "Address", values, disableWhenEmpty, "sender address");
        }

        private void ApplyCategoryCondition(object condition, List<string> values, bool disableWhenEmpty)
        {
            ApplyStringArrayRulePart(condition, "Categories", values, disableWhenEmpty, "category");
        }

        private void ApplyStringArrayRulePart(object part, string propertyName, List<string> values, bool disableWhenEmpty, string label)
        {
            try
            {
                if (values != null && values.Count > 0)
                {
                    SetComProperty(part, propertyName, values.ToArray(), label);
                    SetRulePartEnabled(part, true, label);
                }
                else if (disableWhenEmpty)
                {
                    SetRulePartEnabled(part, false, label);
                }
            }
            finally
            {
                ReleaseComObjectQuietly(part);
            }
        }

        private void ApplyHasAttachmentCondition(object condition, bool? enabled, bool disableWhenEmpty)
        {
            try
            {
                if (enabled.HasValue)
                {
                    if (!enabled.Value)
                        throw new InvalidOperationException("manage_rule upsert: Outlook rules only support the has-attachment condition");
                    SetRulePartEnabled(condition, enabled.Value, "has attachment");
                }
                else if (disableWhenEmpty)
                {
                    SetRulePartEnabled(condition, false, "has attachment");
                }
            }
            finally
            {
                ReleaseComObjectQuietly(condition);
            }
        }

        private void ApplyMoveToFolderAction(object action, string folderPath, bool disableWhenEmpty)
        {
            Outlook.MAPIFolder destFolder = null;
            try
            {
                if (!string.IsNullOrEmpty(folderPath))
                {
                    destFolder = _folderLocator.GetFolderByPath(folderPath);
                    if (destFolder == null)
                        throw new InvalidOperationException("manage_rule upsert: destination folder not found");

                    SetComProperty(action, "Folder", destFolder, "move to folder");
                    SetRulePartEnabled(action, true, "move to folder");
                }
                else if (disableWhenEmpty)
                {
                    SetRulePartEnabled(action, false, "move to folder");
                }
            }
            finally
            {
                ReleaseComObjectQuietly(destFolder);
                ReleaseComObjectQuietly(action);
            }
        }

        private void ApplyAssignToCategoryAction(object action, List<string> categories, bool disableWhenEmpty)
        {
            ApplyStringArrayRulePart(action, "Categories", categories, disableWhenEmpty, "assign category");
        }

        private void ApplyMarkAsTaskAction(object action, bool markAsTask, bool disableWhenEmpty)
        {
            try
            {
                if (markAsTask)
                {
                    SetComProperty(action, "MarkInterval", Outlook.OlMarkInterval.olMarkToday, "mark as task");
                    SetRulePartEnabled(action, true, "mark as task");
                }
                else if (disableWhenEmpty)
                {
                    SetRulePartEnabled(action, false, "mark as task");
                }
            }
            finally
            {
                ReleaseComObjectQuietly(action);
            }
        }

        private void ApplyStopAction(object action, bool stopProcessingMoreRules)
        {
            try
            {
                SetRulePartEnabled(action, stopProcessingMoreRules, "stop processing more rules");
            }
            finally
            {
                ReleaseComObjectQuietly(action);
            }
        }

        private void SetRulePartEnabled(object part, bool enabled, string label)
        {
            SetComProperty(part, "Enabled", enabled, label);
        }

        private void SetComProperty(object target, string propertyName, object value, string label)
        {
            if (target == null)
                throw new InvalidOperationException("manage_rule upsert: missing rule part for " + label);

            try
            {
                var property = target.GetType().GetProperty(propertyName);
                if (property == null)
                    throw new MissingMemberException(target.GetType().FullName, propertyName);
                property.SetValue(target, value);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("manage_rule upsert: failed to apply " + label, ex);
            }
        }

        private void ApplyRuleDelete(Outlook.Rules rules, OutlookCommandRuleRequest req)
        {
            int indexToRemove = -1;

            if (req.OriginalExecutionOrder.HasValue && req.OriginalExecutionOrder.Value > 0
                && req.OriginalExecutionOrder.Value <= rules.Count)
            {
                Outlook.Rule rule = null;
                try
                {
                    rule = rules[req.OriginalExecutionOrder.Value];
                    if (string.IsNullOrEmpty(req.OriginalRuleName) ||
                        string.Equals(rule.Name, req.OriginalRuleName, StringComparison.OrdinalIgnoreCase))
                    {
                        indexToRemove = req.OriginalExecutionOrder.Value;
                    }
                }
                catch { }
                finally
                {
                    ReleaseComObjectQuietly(rule);
                }
            }

            if (indexToRemove < 0 && !string.IsNullOrEmpty(req.OriginalRuleName ?? req.RuleName))
            {
                string name = req.OriginalRuleName ?? req.RuleName;
                for (int i = 1; i <= rules.Count; i++)
                {
                    Outlook.Rule rule = null;
                    try
                    {
                        rule = rules[i];
                        if (string.Equals(rule.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            indexToRemove = i;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        ReleaseComObjectQuietly(rule);
                    }
                }
            }

            if (indexToRemove < 0)
                throw new InvalidOperationException("manage_rule delete: rule not found");

            rules.Remove(indexToRemove);
        }

        private void ApplyRuleSetEnabled(Outlook.Rules rules, OutlookCommandRuleRequest req)
        {
            Outlook.Rule target = null;

            if (req.OriginalExecutionOrder.HasValue && req.OriginalExecutionOrder.Value > 0
                && req.OriginalExecutionOrder.Value <= rules.Count)
            {
                try { target = rules[req.OriginalExecutionOrder.Value]; } catch { }
            }

            if (target == null && !string.IsNullOrEmpty(req.OriginalRuleName ?? req.RuleName))
            {
                string name = req.OriginalRuleName ?? req.RuleName;
                for (int i = 1; i <= rules.Count; i++)
                {
                    Outlook.Rule rule = null;
                    try
                    {
                        rule = rules[i];
                        if (string.Equals(rule.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            target = rule;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (rule != null && !ReferenceEquals(rule, target))
                            ReleaseComObjectQuietly(rule);
                    }
                }
            }

            if (target == null)
                throw new InvalidOperationException("manage_rule set_enabled: rule not found");

            try
            {
                target.Enabled = req.Enabled;
            }
            finally
            {
                ReleaseComObjectQuietly(target);
            }
        }

        private static void ReleaseComObjectQuietly(object obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
