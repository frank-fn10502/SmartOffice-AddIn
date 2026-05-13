using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading.Tasks;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private static readonly HashSet<string> SupportedRuleConditions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Subject",
            "Body",
            "SenderAddress",
            "Category",
            "HasAttachment"
        };

        private static readonly HashSet<string> SupportedRuleActions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "MoveToFolder",
            "AssignToCategory",
            "MarkAsTask",
            "Stop"
        };

        // Read Outlook Rules with best-effort string summaries for conditions/actions.
        public List<OutlookRuleDto> ReadRules()
        {
            var list = new List<OutlookRuleDto>();
            Outlook.Stores stores = null;
            try
            {
                stores = this.Application.Session.Stores;
                if (stores != null)
                {
                    for (int i = 1; i <= stores.Count; i++)
                    {
                        Outlook.Store store = null;
                        try
                        {
                            store = stores[i];
                            ReadRulesFromStore(store, list);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("ReadRules store error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
                        }
                        finally
                        {
                            ReleaseComObjectQuietly(store);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadRules error: " + ex);
            }
            finally
            {
                ReleaseComObjectQuietly(stores);
            }
            return list;
        }

        private void ReadRulesFromStore(Outlook.Store store, List<OutlookRuleDto> list)
        {
            if (store == null || list == null) return;

            Outlook.Rules rules = null;
            string storeId = "";
            try
            {
                try { storeId = store.StoreID ?? ""; } catch { }
                rules = store.GetRules();
                if (rules == null) return;

                for (int i = 1; i <= rules.Count; i++)
                {
                    Outlook.Rule rule = null;
                    try
                    {
                        rule = rules[i];
                        var dto = new OutlookRuleDto
                        {
                            StoreId = storeId,
                            Name = rule.Name ?? "",
                            Enabled = rule.Enabled,
                            ExecutionOrder = rule.ExecutionOrder,
                            RuleType = NormalizeRuleType(rule.RuleType),
                            IsLocalRule = false,
                            Conditions = new List<string>(),
                            Actions = new List<string>(),
                            Exceptions = new List<string>(),
                            CanModifyDefinition = true
                        };

                        bool canModifyDefinition = true;
                        Outlook.RuleConditions conditions = null;
                        Outlook.RuleActions actions = null;
                        Outlook.RuleConditions exceptions = null;
                        try
                        {
                            conditions = rule.Conditions;
                            AddEnabledRuleParts(conditions, dto.Conditions, SupportedRuleConditions, ref canModifyDefinition);

                            actions = rule.Actions;
                            AddEnabledRuleParts(actions, dto.Actions, SupportedRuleActions, ref canModifyDefinition);

                            exceptions = rule.Exceptions;
                            AddEnabledRuleParts(exceptions, dto.Exceptions, null, ref canModifyDefinition);
                        }
                        finally
                        {
                            ReleaseComObjectQuietly(exceptions);
                            ReleaseComObjectQuietly(actions);
                            ReleaseComObjectQuietly(conditions);
                        }

                        dto.CanModifyDefinition = canModifyDefinition;
                        list.Add(dto);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("ReadRules item error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
                    }
                    finally
                    {
                        ReleaseComObjectQuietly(rule);
                    }
                }
            }
            finally
            {
                ReleaseComObjectQuietly(rules);
            }
        }

        private static string NormalizeRuleType(Outlook.OlRuleType ruleType)
        {
            return ruleType == Outlook.OlRuleType.olRuleSend ? "send" : "receive";
        }

        private void AddEnabledRuleParts(
            object parts,
            List<string> summaries,
            HashSet<string> supportedNames,
            ref bool canModifyDefinition)
        {
            if (parts == null || summaries == null) return;

            foreach (var property in parts.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                object value = null;
                try
                {
                    value = property.GetValue(parts);
                    if (value == null) continue;

                    if (!IsRulePartEnabled(value)) continue;

                    var desc = DescribeComObject(value);
                    summaries.Add(string.IsNullOrEmpty(desc)
                        ? property.Name + ": (enabled)"
                        : property.Name + ": " + desc);

                    if (supportedNames == null || !supportedNames.Contains(property.Name))
                        canModifyDefinition = false;
                    if (property.Name.Equals("MoveToFolder", StringComparison.OrdinalIgnoreCase)
                        && !desc.Contains("FolderPath="))
                        canModifyDefinition = false;
                }
                catch { }
                finally
                {
                    ReleaseComObjectQuietly(value);
                }
            }
        }

        private bool IsRulePartEnabled(object rulePart)
        {
            if (rulePart == null) return false;
            try
            {
                var enProp = rulePart.GetType().GetProperty("Enabled");
                return enProp != null && (bool)enProp.GetValue(rulePart);
            }
            catch { return false; }
        }

        // Describe a COM object (condition/action) by reading simple properties (best-effort).
        private string DescribeComObject(object comObj)
        {
            if (comObj == null) return "";
            try
            {
                var t = comObj.GetType();
                var parts = new List<string>();
                foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    try
                    {
                        if (p.Name == "Application" || p.Name == "Parent" || p.Name == "Enabled") continue;
                        object val = null;
                        try
                        {
                            val = p.GetValue(comObj);
                            if (val == null) continue;
                            if (Marshal.IsComObject(val))
                            {
                                var folder = val as Outlook.MAPIFolder;
                                if (folder != null && p.Name == "Folder")
                                {
                                    string folderPath = "";
                                    try { folderPath = folder.FolderPath ?? ""; } catch { }
                                    if (!string.IsNullOrEmpty(folderPath)) parts.Add("FolderPath=" + folderPath);
                                }
                                continue;
                            }
                            if (val is string s)
                            {
                                if (!string.IsNullOrEmpty(s)) parts.Add(p.Name + "=" + s);
                            }
                            else if (val is bool b)
                            {
                                parts.Add(p.Name + "=" + b.ToString());
                            }
                            else if (val is int || val is long || val is double || val is float)
                            {
                                parts.Add(p.Name + "=" + val.ToString());
                            }
                            else if (val is Array array && array.Length > 0)
                            {
                                var arrayParts = new List<string>();
                                foreach (var item in array)
                                {
                                    if (item == null) continue;
                                    var text = item.ToString();
                                    if (!string.IsNullOrEmpty(text)) arrayParts.Add(text);
                                }
                                if (arrayParts.Count > 0) parts.Add(p.Name + "=" + string.Join(", ", arrayParts));
                            }
                        }
                        finally
                        {
                            ReleaseComObjectQuietly(val);
                        }
                    }
                    catch { }
                }
                var result = string.Join("; ", parts);
                if (result.Length > 500) result = result.Substring(0, 500) + "...";
                return result;
            }
            catch { return ""; }
        }

        private static void ReleaseComObjectQuietly(object obj)
        {
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        // ────────────────────────────────────────────────────────────────────────────────
        // manage_rule command handler
        // Supports: upsert, delete, set_enabled
        // Only conditions/actions creatable via Outlook Rules object model are applied.
        // ────────────────────────────────────────────────────────────────────────────────
        private async Task HandleManageRuleAsync(OutlookCommand cmd)
        {
            var req = cmd.RuleRequest;
            if (req == null || string.IsNullOrEmpty(req.Operation))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "manage_rule failed: operation is required");
                return;
            }

            Exception ruleEx = null;
            _chatPane.Invoke((Action)(() =>
            {
                try { ApplyManageRule(req); }
                catch (Exception ex) { ruleEx = ex; }
            }));

            if (ruleEx != null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "manage_rule error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ruleEx));
                return;
            }

            // Push updated snapshot
            try
            {
                List<OutlookRuleDto> updatedRules = null;
                _chatPane.Invoke((Action)(() => { updatedRules = ReadRules(); }));
                await _signalRClient.PushRulesAsync(updatedRules ?? new List<OutlookRuleDto>());
            }
            catch { }

            await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                $"manage_rule {req.Operation} completed");
        }

        /// <summary>
        /// Applies the manage_rule operation on the UI thread.
        /// Throws on failure; caller translates to ReportCommandResult(success=false).
        /// </summary>
        private void ApplyManageRule(OutlookCommandRuleRequest req)
        {
            Outlook.Store store = null;
            Outlook.Rules rules = null;
            try
            {
                // Resolve store: use provided storeId or fall back to default store
                if (!string.IsNullOrEmpty(req.StoreId))
                {
                    var stores = this.Application.Session.Stores;
                    try
                    {
                        foreach (Outlook.Store s in stores)
                        {
                            string sid = "";
                            try { sid = s.StoreID ?? ""; } catch { }
                            if (string.Equals(sid, req.StoreId, StringComparison.OrdinalIgnoreCase))
                            {
                                store = s;
                                break;
                            }
                            else
                            {
                                try { Marshal.ReleaseComObject(s); } catch { }
                            }
                        }
                    }
                    finally { try { Marshal.ReleaseComObject(stores); } catch { } }

                    if (store == null)
                        throw new InvalidOperationException("manage_rule: requested Outlook store was not found");
                }

                if (store == null)
                    store = this.Application.Session.DefaultStore;

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

                // Persist changes; pass false to avoid showing UI dialog
                rules.Save(false);
            }
            finally
            {
                if (rules != null) try { Marshal.ReleaseComObject(rules); } catch { }
                if (store != null) try { Marshal.ReleaseComObject(store); } catch { }
            }
        }

        private void ApplyRuleUpsert(Outlook.Rules rules, OutlookCommandRuleRequest req)
        {
            // Try to find existing rule using originalExecutionOrder first, then originalRuleName
            Outlook.Rule existingRule = null;
            if (req.OriginalExecutionOrder.HasValue && req.OriginalExecutionOrder.Value > 0)
            {
                try
                {
                    var r = rules[req.OriginalExecutionOrder.Value];
                    if (r != null) existingRule = r;
                }
                catch { }
            }

            if (existingRule == null && !string.IsNullOrEmpty(req.OriginalRuleName))
            {
                try
                {
                    for (int i = 1; i <= rules.Count; i++)
                    {
                        Outlook.Rule r = null;
                        try
                        {
                            r = rules[i];
                            if (string.Equals(r.Name, req.OriginalRuleName, StringComparison.OrdinalIgnoreCase))
                            {
                                existingRule = r;
                                break;
                            }
                        }
                        catch { }
                        finally { if (r != null && !ReferenceEquals(r, existingRule)) try { Marshal.ReleaseComObject(r); } catch { } }
                    }
                }
                catch { }
            }

            Outlook.Rule rule = existingRule;
            try
            {
                if (rule == null)
                {
                    var ruleName = (req.RuleName ?? "").Trim();
                    if (string.IsNullOrEmpty(ruleName))
                        throw new InvalidOperationException("manage_rule upsert: rule name is required");

                    var ruleType = string.Equals(req.RuleType, "send", StringComparison.OrdinalIgnoreCase)
                        ? Outlook.OlRuleType.olRuleSend
                        : Outlook.OlRuleType.olRuleReceive;
                    rule = rules.Create(ruleName, ruleType);
                }
                else if (!string.IsNullOrEmpty(req.RuleName))
                {
                    rule.Name = req.RuleName;
                }

                rule.Enabled = req.Enabled;
                if (req.ExecutionOrder.HasValue && req.ExecutionOrder.Value > 0)
                {
                    rule.ExecutionOrder = req.ExecutionOrder.Value;
                }

                if (req.Conditions != null)
                {
                    Outlook.RuleConditions conditions = null;
                    try
                    {
                        conditions = rule.Conditions;
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
                        actions = rule.Actions;
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
                ReleaseComObjectQuietly(rule);
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
                    SetRulePartEnabled(condition, false, "has attachment");
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
                    destFolder = GetFolderByPath(folderPath);
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

            // Prefer originalExecutionOrder for precise targeting
            if (req.OriginalExecutionOrder.HasValue && req.OriginalExecutionOrder.Value > 0
                && req.OriginalExecutionOrder.Value <= rules.Count)
            {
                Outlook.Rule r = null;
                try
                {
                    r = rules[req.OriginalExecutionOrder.Value];
                    // Confirm name matches if provided
                    if (string.IsNullOrEmpty(req.OriginalRuleName) ||
                        string.Equals(r.Name, req.OriginalRuleName, StringComparison.OrdinalIgnoreCase))
                    {
                        indexToRemove = req.OriginalExecutionOrder.Value;
                    }
                }
                catch { }
                finally { if (r != null) try { Marshal.ReleaseComObject(r); } catch { } }
            }

            // Fall back to name search
            if (indexToRemove < 0 && !string.IsNullOrEmpty(req.OriginalRuleName ?? req.RuleName))
            {
                string name = req.OriginalRuleName ?? req.RuleName;
                for (int i = 1; i <= rules.Count; i++)
                {
                    Outlook.Rule r = null;
                    try
                    {
                        r = rules[i];
                        if (string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            indexToRemove = i;
                            break;
                        }
                    }
                    catch { }
                    finally { if (r != null) try { Marshal.ReleaseComObject(r); } catch { } }
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
                    Outlook.Rule r = null;
                    try
                    {
                        r = rules[i];
                        if (string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            target = r;
                            break;
                        }
                    }
                    catch { }
                    finally { if (r != null && !ReferenceEquals(r, target)) try { Marshal.ReleaseComObject(r); } catch { } }
                }
            }

            if (target == null)
                throw new InvalidOperationException("manage_rule set_enabled: rule not found");

            try
            {
                target.Enabled = req.Enabled;
            }
            finally { try { Marshal.ReleaseComObject(target); } catch { } }
        }
    }
}
