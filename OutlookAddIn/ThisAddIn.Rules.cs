using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        // Read Outlook Rules with best-effort string summaries for conditions/actions
        public List<OutlookRuleDto> ReadRules()
        {
            var list = new List<OutlookRuleDto>();
            try
            {
                // Try to get Rules collection from default store
                Outlook.Rules rules = null;
                try
                {
                    var store = this.Application.Session.DefaultStore;
                    rules = store?.GetRules();
                    if (rules == null)
                    {
                        try { Marshal.ReleaseComObject(store); } catch { }
                        return list;
                    }

                    int order = 1;
                    foreach (Outlook.Rule r in rules)
                    {
                        try
                        {
                            var dto = new OutlookRuleDto
                            {
                                Name = r.Name ?? "",
                                Enabled = r.Enabled,
                                ExecutionOrder = order++,
                                RuleType = r.RuleType.ToString(),
                                Conditions = new List<string>(),
                                Actions = new List<string>(),
                                Exceptions = new List<string>(),
                                CanModifyDefinition = true
                            };

                            // Best-effort: describe known condition/action/exception sub-objects
                            try
                            {
                                var conds = r.Conditions;
                                if (conds != null)
                                {
                                    // Iterate properties of Conditions and describe enabled ones
                                    foreach (var p in conds.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                                    {
                                        try
                                        {
                                            var val = p.GetValue(conds);
                                            if (val == null) continue;
                                            // Many condition properties have Enabled bool; inspect that
                                            bool enabled = false;
                                            try
                                            {
                                                var enProp = val.GetType().GetProperty("Enabled");
                                                if (enProp != null)
                                                {
                                                    enabled = (bool)enProp.GetValue(val);
                                                }
                                            }
                                            catch { }

                                            if (enabled)
                                            {
                                                var desc = DescribeComObject(val);
                                                if (!string.IsNullOrEmpty(desc)) dto.Conditions.Add(p.Name + ": " + desc);
                                                else dto.Conditions.Add(p.Name + ": (enabled)");

                                                try { Marshal.ReleaseComObject(val); } catch { }
                                            }
                                        }
                                        catch { }
                                    }
                                    try { Marshal.ReleaseComObject(conds); } catch { }
                                }
                            }
                            catch { }

                            try
                            {
                                var acts = r.Actions;
                                if (acts != null)
                                {
                                    foreach (var p in acts.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                                    {
                                        try
                                        {
                                            var val = p.GetValue(acts);
                                            if (val == null) continue;
                                            bool enabled = false;
                                            try
                                            {
                                                var enProp = val.GetType().GetProperty("Enabled");
                                                if (enProp != null)
                                                {
                                                    enabled = (bool)enProp.GetValue(val);
                                                }
                                            }
                                            catch { }

                                            if (enabled)
                                            {
                                                var desc = DescribeComObject(val);
                                                if (!string.IsNullOrEmpty(desc)) dto.Actions.Add(p.Name + ": " + desc);
                                                else dto.Actions.Add(p.Name + ": (enabled)");

                                                try { Marshal.ReleaseComObject(val); } catch { }
                                            }
                                        }
                                        catch { }
                                    }
                                    try { Marshal.ReleaseComObject(acts); } catch { }
                                }
                            }
                            catch { }

                            try
                            {
                                var excs = r.Exceptions;
                                if (excs != null)
                                {
                                    foreach (var p in excs.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                                    {
                                        try
                                        {
                                            var val = p.GetValue(excs);
                                            if (val == null) continue;
                                            bool enabled = false;
                                            try
                                            {
                                                var enProp = val.GetType().GetProperty("Enabled");
                                                if (enProp != null)
                                                {
                                                    enabled = (bool)enProp.GetValue(val);
                                                }
                                            }
                                            catch { }

                                            if (enabled)
                                            {
                                                var desc = DescribeComObject(val);
                                                if (!string.IsNullOrEmpty(desc)) dto.Exceptions.Add(p.Name + ": " + desc);
                                                else dto.Exceptions.Add(p.Name + ": (enabled)");

                                                try { Marshal.ReleaseComObject(val); } catch { }
                                            }
                                        }
                                        catch { }
                                    }
                                    try { Marshal.ReleaseComObject(excs); } catch { }
                                }
                            }
                            catch { }

                            list.Add(dto);
                        }
                        catch { }
                        finally
                        {
                            try { Marshal.ReleaseComObject(r); } catch { }
                        }
                    }
                }
                finally
                {
                    try { if (rules != null) Marshal.ReleaseComObject(rules); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadRules error: " + ex);
            }
            return list;
        }

        // Describe a COM object (condition/action) by reading simple properties (best-effort)
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
                        var val = p.GetValue(comObj);
                        if (val == null) continue;
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
                        else
                        {
                            // For some properties (collections) skip to avoid verbose output
                            var sval = val.ToString();
                            if (!string.IsNullOrEmpty(sval) && sval.Length < 200)
                                parts.Add(p.Name + "=" + sval);
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
                    "manage_rule error: " + SanitizeExceptionForLog(ruleEx));
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
            if (rule == null)
            {
                // Create new rule
                var ruleType = string.Equals(req.RuleType, "send", StringComparison.OrdinalIgnoreCase)
                    ? Outlook.OlRuleType.olRuleSend
                    : Outlook.OlRuleType.olRuleReceive;
                rule = rules.Create(req.RuleName ?? "New Rule", ruleType);
            }
            else
            {
                // Update existing rule name
                if (!string.IsNullOrEmpty(req.RuleName))
                    rule.Name = req.RuleName;
            }

            rule.Enabled = req.Enabled;

            // Apply supported conditions
            if (req.Conditions != null)
            {
                // SubjectContains (Outlook: Conditions.Subject)
                try
                {
                    var cond = rule.Conditions.Subject;
                    if (req.Conditions.SubjectContains != null && req.Conditions.SubjectContains.Count > 0)
                    {
                        cond.Text = req.Conditions.SubjectContains.ToArray();
                        cond.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        cond.Enabled = false;
                    }
                    Marshal.ReleaseComObject(cond);
                }
                catch { }

                // BodyContains (Outlook: Conditions.Body)
                try
                {
                    var cond = rule.Conditions.Body;
                    if (req.Conditions.BodyContains != null && req.Conditions.BodyContains.Count > 0)
                    {
                        cond.Text = req.Conditions.BodyContains.ToArray();
                        cond.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        cond.Enabled = false;
                    }
                    Marshal.ReleaseComObject(cond);
                }
                catch { }

                // SenderAddressContains
                try
                {
                    var cond = rule.Conditions.SenderAddress;
                    if (req.Conditions.SenderAddressContains != null && req.Conditions.SenderAddressContains.Count > 0)
                    {
                        cond.Address = req.Conditions.SenderAddressContains.ToArray();
                        cond.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        cond.Enabled = false;
                    }
                    Marshal.ReleaseComObject(cond);
                }
                catch { }

                // Categories (Outlook: Conditions.Category → CategoryRuleCondition.Categories)
                try
                {
                    var cond = rule.Conditions.Category;
                    if (req.Conditions.Categories != null && req.Conditions.Categories.Count > 0)
                    {
                        cond.Categories = req.Conditions.Categories.ToArray();
                        cond.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        cond.Enabled = false;
                    }
                    Marshal.ReleaseComObject(cond);
                }
                catch { }

                // HasAttachment
                try
                {
                    var cond = rule.Conditions.HasAttachment;
                    if (req.Conditions.HasAttachment.HasValue)
                        cond.Enabled = req.Conditions.HasAttachment.Value;
                    else if (existingRule != null)
                        cond.Enabled = false;
                    Marshal.ReleaseComObject(cond);
                }
                catch { }
            }

            // Apply supported actions
            if (req.Actions != null)
            {
                // MoveToFolder
                try
                {
                    var act = rule.Actions.MoveToFolder;
                    if (!string.IsNullOrEmpty(req.Actions.MoveToFolderPath))
                    {
                        var destFolder = GetFolderByPath(req.Actions.MoveToFolderPath);
                        if (destFolder != null)
                        {
                            act.Folder = destFolder;
                            act.Enabled = true;
                            Marshal.ReleaseComObject(destFolder);
                        }
                    }
                    else if (existingRule != null)
                    {
                        act.Enabled = false;
                    }
                    Marshal.ReleaseComObject(act);
                }
                catch { }

                // AssignCategories
                try
                {
                    var act = rule.Actions.AssignToCategory;
                    if (req.Actions.AssignCategories != null && req.Actions.AssignCategories.Count > 0)
                    {
                        act.Categories = req.Actions.AssignCategories.ToArray();
                        act.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        act.Enabled = false;
                    }
                    Marshal.ReleaseComObject(act);
                }
                catch { }

                // MarkAsTask
                try
                {
                    var act = rule.Actions.MarkAsTask;
                    if (req.Actions.MarkAsTask)
                    {
                        act.MarkInterval = Outlook.OlMarkInterval.olMarkToday;
                        act.Enabled = true;
                    }
                    else if (existingRule != null)
                    {
                        act.Enabled = false;
                    }
                    Marshal.ReleaseComObject(act);
                }
                catch { }

                // StopProcessingMoreRules
                try
                {
                    var act = rule.Actions.Stop;
                    act.Enabled = req.Actions.StopProcessingMoreRules;
                    Marshal.ReleaseComObject(act);
                }
                catch { }
            }

            if (existingRule != null)
                try { Marshal.ReleaseComObject(existingRule); } catch { }
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
