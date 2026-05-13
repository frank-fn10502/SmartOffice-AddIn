using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAddIn.Infrastructure.Diagnostics;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Rules
{
    internal sealed class OutlookRuleReader
    {
        private static readonly HashSet<string> SupportedRuleConditions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Subject",
            "Body",
            "BodyOrSubject",
            "MessageHeader",
            "SenderAddress",
            "RecipientAddress",
            "Category",
            "HasAttachment",
            "Importance",
            "ToMe",
            "ToOrCc",
            "OnlyToMe",
            "MeetingInviteOrUpdate"
        };

        private static readonly HashSet<string> SupportedRuleActions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "MoveToFolder",
            "CopyToFolder",
            "AssignToCategory",
            "ClearCategories",
            "MarkAsTask",
            "Delete",
            "DesktopAlert",
            "Stop"
        };

        private readonly Outlook.Application _application;

        public OutlookRuleReader(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public List<OutlookRuleDto> ReadRules()
        {
            var list = new List<OutlookRuleDto>();
            Outlook.Stores stores = null;

            try
            {
                stores = _application.Session.Stores;
                if (stores == null)
                    return list;

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
                        System.Diagnostics.Debug.WriteLine("ReadRules store error: " + SensitiveLogSanitizer.Sanitize(ex));
                    }
                    finally
                    {
                        Release(store);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadRules error: " + ex);
            }
            finally
            {
                Release(stores);
            }

            return list;
        }

        private static void ReadRulesFromStore(Outlook.Store store, List<OutlookRuleDto> list)
        {
            if (store == null || list == null)
                return;

            Outlook.Rules rules = null;
            var storeId = "";

            try
            {
                try { storeId = store.StoreID ?? ""; } catch { }
                rules = store.GetRules();
                if (rules == null)
                    return;

                for (int i = 1; i <= rules.Count; i++)
                {
                    Outlook.Rule rule = null;
                    try
                    {
                        rule = rules[i];
                        var dto = BuildRuleDto(storeId, rule);
                        list.Add(dto);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("ReadRules item error: " + SensitiveLogSanitizer.Sanitize(ex));
                    }
                    finally
                    {
                        Release(rule);
                    }
                }
            }
            finally
            {
                Release(rules);
            }
        }

        private static OutlookRuleDto BuildRuleDto(string storeId, Outlook.Rule rule)
        {
            var dto = new OutlookRuleDto
            {
                StoreId = storeId,
                Name = ReadString(() => rule.Name),
                Enabled = ReadBool(() => rule.Enabled),
                ExecutionOrder = ReadInt(() => rule.ExecutionOrder),
                RuleType = ReadRuleType(rule),
                IsLocalRule = false,
                Conditions = new List<string>(),
                Actions = new List<string>(),
                Exceptions = new List<string>(),
                CanModifyDefinition = true
            };

            var canModifyDefinition = true;
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
                Release(exceptions);
                Release(actions);
                Release(conditions);
            }

            dto.CanModifyDefinition = canModifyDefinition;
            return dto;
        }

        private static string NormalizeRuleType(Outlook.OlRuleType ruleType)
        {
            return ruleType == Outlook.OlRuleType.olRuleSend ? "send" : "receive";
        }

        private static string ReadRuleType(Outlook.Rule rule)
        {
            try { return NormalizeRuleType(rule.RuleType); } catch { return "receive"; }
        }

        private static void AddEnabledRuleParts(
            object parts,
            List<string> summaries,
            HashSet<string> supportedNames,
            ref bool canModifyDefinition)
        {
            if (parts == null || summaries == null)
                return;

            foreach (var property in parts.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                object value = null;
                try
                {
                    value = property.GetValue(parts);
                    if (value == null)
                        continue;

                    if (!IsRulePartEnabled(value))
                        continue;

                    var description = DescribeComObject(value);
                    summaries.Add(string.IsNullOrEmpty(description)
                        ? property.Name + ": (enabled)"
                        : property.Name + ": " + description);

                    if (supportedNames == null || !supportedNames.Contains(property.Name))
                        canModifyDefinition = false;
                    if (property.Name.Equals("MoveToFolder", StringComparison.OrdinalIgnoreCase)
                        && !description.Contains("FolderPath="))
                        canModifyDefinition = false;
                }
                catch { }
                finally
                {
                    Release(value);
                }
            }
        }

        private static bool IsRulePartEnabled(object rulePart)
        {
            if (rulePart == null)
                return false;

            try
            {
                var enabledProperty = rulePart.GetType().GetProperty("Enabled");
                return enabledProperty != null && (bool)enabledProperty.GetValue(rulePart);
            }
            catch { return false; }
        }

        private static string DescribeComObject(object comObj)
        {
            if (comObj == null)
                return "";

            try
            {
                var type = comObj.GetType();
                var parts = new List<string>();
                foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    try
                    {
                        if (property.Name == "Application" || property.Name == "Parent" || property.Name == "Enabled")
                            continue;

                        object value = null;
                        try
                        {
                            value = property.GetValue(comObj);
                            if (value == null)
                                continue;

                            if (Marshal.IsComObject(value))
                            {
                                var folder = value as Outlook.MAPIFolder;
                                if (folder != null && property.Name == "Folder")
                                {
                                    var folderPath = "";
                                    try { folderPath = folder.FolderPath ?? ""; } catch { }
                                    if (!string.IsNullOrEmpty(folderPath))
                                        parts.Add("FolderPath=" + folderPath);
                                }
                                continue;
                            }

                            if (value is string s)
                            {
                                if (!string.IsNullOrEmpty(s))
                                    parts.Add(property.Name + "=" + s);
                            }
                            else if (value is bool b)
                            {
                                parts.Add(property.Name + "=" + b.ToString());
                            }
                            else if (value is int || value is long || value is double || value is float)
                            {
                                parts.Add(property.Name + "=" + value.ToString());
                            }
                            else if (value is Array array && array.Length > 0)
                            {
                                var arrayParts = new List<string>();
                                foreach (var item in array)
                                {
                                    if (item == null)
                                        continue;

                                    var text = item.ToString();
                                    if (!string.IsNullOrEmpty(text))
                                        arrayParts.Add(text);
                                }

                                if (arrayParts.Count > 0)
                                    parts.Add(property.Name + "=" + string.Join(", ", arrayParts));
                            }
                        }
                        finally
                        {
                            Release(value);
                        }
                    }
                    catch { }
                }

                var result = string.Join("; ", parts);
                return result.Length > 500 ? result.Substring(0, 500) + "..." : result;
            }
            catch { return ""; }
        }

        private static string ReadString(Func<string> read)
        {
            try { return read() ?? ""; } catch { return ""; }
        }

        private static bool ReadBool(Func<bool> read)
        {
            try { return read(); } catch { return false; }
        }

        private static int ReadInt(Func<int> read)
        {
            try { return read(); } catch { return 0; }
        }

        private static void Release(object obj)
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
