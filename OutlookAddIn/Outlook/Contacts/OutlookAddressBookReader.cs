using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Contacts
{
    internal sealed class OutlookAddressBookReader
    {
        private readonly Outlook.Application _application;
        private const int UiPumpInterval = 10;

        public OutlookAddressBookReader(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public List<AddressBookContactDto> ReadAddressBook(AddressBookSyncRequest request, Action<List<AddressBookContactDto>> publishSnapshot = null)
        {
            request = request ?? new AddressBookSyncRequest();
            var maxContacts = Clamp(request.MaxContacts, 1, 5000, 1000);
            var maxAddressEntriesPerList = Clamp(request.MaxAddressEntriesPerList, 1, 2000, 500);
            var maxGroupMembers = request.MaxGroupMembers < 0 ? 50 : Math.Min(request.MaxGroupMembers, 500);
            var maxGroupDepth = request.MaxGroupDepth < 0 ? 1 : Math.Min(request.MaxGroupDepth, 3);
            var groupMemberReadBudget = Math.Min(maxContacts, 1000);
            var contacts = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);
            var publishThreshold = 50;
            var lastPublishedCount = 0;

            if (request.IncludeOutlookContacts)
                ReadDefaultContactsFolder(contacts, maxContacts, publishSnapshot, publishThreshold, ref lastPublishedCount);

            if (request.IncludeAddressLists && contacts.Count < maxContacts)
                ReadAddressLists(contacts, maxContacts, maxAddressEntriesPerList, maxGroupMembers, maxGroupDepth, ref groupMemberReadBudget, publishSnapshot, publishThreshold, ref lastPublishedCount);

            var result = contacts.Values
                .OrderBy(item => item.DisplayName)
                .ThenBy(item => item.SmtpAddress)
                .Take(maxContacts)
                .ToList();
            publishSnapshot?.Invoke(CloneContacts(result));
            return result;
        }

        private void ReadDefaultContactsFolder(
            Dictionary<string, AddressBookContactDto> contacts,
            int maxContacts,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            ref int lastPublishedCount)
        {
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;

            try
            {
                folder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                items = folder?.Items;
                if (items == null) return;

                var processedSincePump = 0;
                for (var i = 1; i <= items.Count && contacts.Count < maxContacts; i++)
                {
                    object item = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        item = items[i];
                        contact = item as Outlook.ContactItem;
                        if (contact == null) continue;

                        AddContactItem(contacts, contact, 1);
                        AddContactItem(contacts, contact, 2);
                        AddContactItem(contacts, contact, 3);
                        PublishIfNeeded(contacts, publishSnapshot, publishThreshold, ref lastPublishedCount);
                        PumpOutlookUi(ref processedSincePump);
                    }
                    catch { }
                    finally
                    {
                        Release(contact);
                        if (!ReferenceEquals(item, contact)) Release(item);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadDefaultContactsFolder error: " + ex);
            }
            finally
            {
                Release(items);
                Release(folder);
            }
        }

        private void ReadAddressLists(
            Dictionary<string, AddressBookContactDto> contacts,
            int maxContacts,
            int maxAddressEntriesPerList,
            int maxGroupMembers,
            int maxGroupDepth,
            ref int groupMemberReadBudget,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            ref int lastPublishedCount)
        {
            Outlook.AddressLists lists = null;

            try
            {
                lists = _application.Session.AddressLists;
                if (lists == null) return;

                for (var listIndex = 1; listIndex <= lists.Count && contacts.Count < maxContacts; listIndex++)
                {
                    Outlook.AddressList list = null;
                    Outlook.AddressEntries entries = null;
                    try
                    {
                        list = lists[listIndex];
                        if (!IsSupportedAddressList(list)) continue;

                        entries = list.AddressEntries;
                        if (entries == null) continue;

                        var entryLimit = Math.Min(entries.Count, maxAddressEntriesPerList);
                        var processedSincePump = 0;
                        for (var entryIndex = 1; entryIndex <= entryLimit && contacts.Count < maxContacts; entryIndex++)
                        {
                            Outlook.AddressEntry entry = null;
                            try
                            {
                                entry = entries[entryIndex];
                                AddAddressEntry(contacts, entry, list, maxGroupMembers, maxGroupDepth, ref groupMemberReadBudget);
                                PublishIfNeeded(contacts, publishSnapshot, publishThreshold, ref lastPublishedCount);
                                PumpOutlookUi(ref processedSincePump);
                            }
                            catch { }
                            finally
                            {
                                Release(entry);
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        Release(entries);
                        Release(list);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadAddressLists error: " + ex);
            }
            finally
            {
                Release(lists);
            }
        }

        private static bool IsSupportedAddressList(Outlook.AddressList list)
        {
            try
            {
                var type = list.AddressListType;
                return type == Outlook.OlAddressListType.olOutlookAddressList
                    || type == Outlook.OlAddressListType.olExchangeGlobalAddressList
                    || type == Outlook.OlAddressListType.olCustomAddressList;
            }
            catch
            {
                return false;
            }
        }

        private static void AddContactItem(Dictionary<string, AddressBookContactDto> contacts, Outlook.ContactItem item, int emailSlot)
        {
            var email = ReadContactEmail(item, emailSlot);
            if (string.IsNullOrWhiteSpace(email)) return;

            Upsert(contacts, new AddressBookContactDto
            {
                Id = ReadString(() => item.EntryID) + ":email" + emailSlot,
                DisplayName = ReadString(() => item.FullName),
                SmtpAddress = email,
                RawAddress = email,
                AddressType = ReadContactEmailType(item, emailSlot),
                EntryUserType = "olOutlookContactAddressEntry",
                Source = "outlook_contacts",
                CompanyName = ReadString(() => item.CompanyName),
                JobTitle = ReadString(() => item.JobTitle),
                Department = ReadString(() => item.Department),
                OfficeLocation = ReadString(() => item.OfficeLocation),
                BusinessTelephoneNumber = ReadString(() => item.BusinessTelephoneNumber),
                MobileTelephoneNumber = ReadString(() => item.MobileTelephoneNumber),
            });
        }

        private static void AddAddressEntry(
            Dictionary<string, AddressBookContactDto> contacts,
            Outlook.AddressEntry entry,
            Outlook.AddressList list,
            int maxGroupMembers,
            int maxGroupDepth,
            ref int groupMemberReadBudget)
        {
            if (entry == null) return;

            var dto = new AddressBookContactDto
            {
                Id = ReadString(() => entry.ID),
                DisplayName = ReadString(() => entry.Name),
                SmtpAddress = ReadString(() => entry.Address),
                RawAddress = ReadString(() => entry.Address),
                AddressType = ReadString(() => entry.Type),
                EntryUserType = ReadString(() => entry.AddressEntryUserType.ToString()),
                Source = AddressListSource(list),
            };

            Outlook.ExchangeUser exchangeUser = null;
            Outlook.ContactItem contactItem = null;
            Outlook.ExchangeDistributionList distributionList = null;
            try
            {
                exchangeUser = entry.GetExchangeUser();
                if (exchangeUser != null)
                {
                    dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => exchangeUser.Name));
                    dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => exchangeUser.PrimarySmtpAddress));
                    dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                    dto.CompanyName = ReadString(() => exchangeUser.CompanyName);
                    dto.JobTitle = ReadString(() => exchangeUser.JobTitle);
                    dto.Department = ReadString(() => exchangeUser.Department);
                    dto.OfficeLocation = ReadString(() => exchangeUser.OfficeLocation);
                    dto.BusinessTelephoneNumber = ReadString(() => exchangeUser.BusinessTelephoneNumber);
                    dto.MobileTelephoneNumber = ReadString(() => exchangeUser.MobileTelephoneNumber);
                }

                contactItem = entry.GetContact();
                if (contactItem != null)
                {
                    dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => contactItem.FullName));
                    dto.CompanyName = Prefer(dto.CompanyName, ReadString(() => contactItem.CompanyName));
                    dto.JobTitle = Prefer(dto.JobTitle, ReadString(() => contactItem.JobTitle));
                    dto.Department = Prefer(dto.Department, ReadString(() => contactItem.Department));
                    dto.OfficeLocation = Prefer(dto.OfficeLocation, ReadString(() => contactItem.OfficeLocation));
                    dto.BusinessTelephoneNumber = Prefer(dto.BusinessTelephoneNumber, ReadString(() => contactItem.BusinessTelephoneNumber));
                    dto.MobileTelephoneNumber = Prefer(dto.MobileTelephoneNumber, ReadString(() => contactItem.MobileTelephoneNumber));
                }

                distributionList = entry.GetExchangeDistributionList();
                if (distributionList != null)
                {
                    dto.IsGroup = true;
                    dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => distributionList.Name));
                    dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => distributionList.PrimarySmtpAddress));
                    dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                    ReadDistributionListMembers(distributionList, dto, maxGroupMembers, maxGroupDepth, 0, new HashSet<string>(StringComparer.OrdinalIgnoreCase), ref groupMemberReadBudget);
                }
            }
            catch { }
            finally
            {
                Release(distributionList);
                Release(contactItem);
                Release(exchangeUser);
            }

            if (!string.IsNullOrWhiteSpace(dto.SmtpAddress) || !string.IsNullOrWhiteSpace(dto.DisplayName))
                Upsert(contacts, dto);
        }

        private static void ReadDistributionListMembers(
            Outlook.ExchangeDistributionList distributionList,
            AddressBookContactDto dto,
            int maxGroupMembers,
            int maxGroupDepth,
            int depth,
            HashSet<string> visitedGroups,
            ref int groupMemberReadBudget)
        {
            if (distributionList == null || maxGroupMembers <= 0 || depth > maxGroupDepth || groupMemberReadBudget <= 0) return;
            var groupKey = ReadString(() => distributionList.PrimarySmtpAddress);
            if (!string.IsNullOrWhiteSpace(groupKey) && !visitedGroups.Add(groupKey)) return;

            Outlook.AddressEntries members = null;
            try
            {
                members = distributionList.GetExchangeDistributionListMembers();
                if (members == null) return;
                var limit = Math.Min(Math.Min(members.Count, maxGroupMembers - dto.MemberSmtpAddresses.Count), groupMemberReadBudget);
                var processedSincePump = 0;
                for (var i = 1; i <= limit; i++)
                {
                    groupMemberReadBudget--;
                    Outlook.AddressEntry member = null;
                    Outlook.ExchangeDistributionList nested = null;
                    try
                    {
                        member = members[i];
                        var smtp = SmtpFromAddressEntry(member);
                        if (string.IsNullOrWhiteSpace(smtp)) smtp = ReadString(() => member.Address);
                        if (!string.IsNullOrWhiteSpace(smtp) && !dto.MemberSmtpAddresses.Contains(smtp, StringComparer.OrdinalIgnoreCase))
                            dto.MemberSmtpAddresses.Add(smtp);

                        nested = member.GetExchangeDistributionList();
                        if (nested != null && !string.IsNullOrWhiteSpace(smtp))
                        {
                            dto.MemberGroupSmtpAddresses.Add(smtp);
                            ReadDistributionListMembers(nested, dto, maxGroupMembers, maxGroupDepth, depth + 1, visitedGroups, ref groupMemberReadBudget);
                        }
                        PumpOutlookUi(ref processedSincePump);
                    }
                    catch { }
                    finally
                    {
                        Release(nested);
                        Release(member);
                    }
                }
                dto.MemberCount = Math.Max(dto.MemberCount, members.Count);
            }
            catch { }
            finally
            {
                Release(members);
            }
        }

        private static void PumpOutlookUi(ref int processedSincePump)
        {
            processedSincePump++;
            if (processedSincePump < UiPumpInterval) return;
            processedSincePump = 0;
            try
            {
                Application.DoEvents();
            }
            catch { }
        }

        private static void Upsert(Dictionary<string, AddressBookContactDto> contacts, AddressBookContactDto dto)
        {
            var key = Normalize(dto.SmtpAddress);
            if (string.IsNullOrWhiteSpace(key))
                key = Normalize(dto.DisplayName);
            if (string.IsNullOrWhiteSpace(key)) return;

            dto.Domain = EmailDomain(dto.SmtpAddress);
            dto.IsKnown = true;
            dto.Sources = new List<string> { dto.Source };
            dto.RelationKinds = new List<string> { "address_book" };
            dto.MemberSmtpAddresses = dto.MemberSmtpAddresses.Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();
            dto.MemberGroupSmtpAddresses = dto.MemberGroupSmtpAddresses.Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();

            AddressBookContactDto current;
            if (!contacts.TryGetValue(key, out current))
            {
                contacts[key] = dto;
                return;
            }

            current.DisplayName = Prefer(current.DisplayName, dto.DisplayName);
            current.SmtpAddress = Prefer(current.SmtpAddress, dto.SmtpAddress);
            current.RawAddress = Prefer(current.RawAddress, dto.RawAddress);
            current.AddressType = Prefer(current.AddressType, dto.AddressType);
            current.EntryUserType = Prefer(current.EntryUserType, dto.EntryUserType);
            current.CompanyName = Prefer(current.CompanyName, dto.CompanyName);
            current.JobTitle = Prefer(current.JobTitle, dto.JobTitle);
            current.Department = Prefer(current.Department, dto.Department);
            current.OfficeLocation = Prefer(current.OfficeLocation, dto.OfficeLocation);
            current.BusinessTelephoneNumber = Prefer(current.BusinessTelephoneNumber, dto.BusinessTelephoneNumber);
            current.MobileTelephoneNumber = Prefer(current.MobileTelephoneNumber, dto.MobileTelephoneNumber);
            current.IsGroup = current.IsGroup || dto.IsGroup;
            current.MemberCount = Math.Max(current.MemberCount, dto.MemberCount);
            current.MemberSmtpAddresses = current.MemberSmtpAddresses.Concat(dto.MemberSmtpAddresses).Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();
            current.MemberGroupSmtpAddresses = current.MemberGroupSmtpAddresses.Concat(dto.MemberGroupSmtpAddresses).Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();
            if (!current.Sources.Contains(dto.Source)) current.Sources.Add(dto.Source);
        }

        private static void PublishIfNeeded(
            Dictionary<string, AddressBookContactDto> contacts,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            ref int lastPublishedCount)
        {
            if (publishSnapshot == null) return;
            if (contacts.Count - lastPublishedCount < publishThreshold) return;
            lastPublishedCount = contacts.Count;
            publishSnapshot(CloneContacts(contacts.Values));
        }

        private static List<AddressBookContactDto> CloneContacts(IEnumerable<AddressBookContactDto> contacts)
        {
            return contacts
                .Select(CloneContact)
                .OrderBy(item => item.DisplayName)
                .ThenBy(item => item.SmtpAddress)
                .ToList();
        }

        private static AddressBookContactDto CloneContact(AddressBookContactDto contact)
        {
            return new AddressBookContactDto
            {
                Id = contact.Id,
                DisplayName = contact.DisplayName,
                SmtpAddress = contact.SmtpAddress,
                RawAddress = contact.RawAddress,
                AddressType = contact.AddressType,
                EntryUserType = contact.EntryUserType,
                Source = contact.Source,
                Sources = new List<string>(contact.Sources ?? new List<string>()),
                CompanyName = contact.CompanyName,
                JobTitle = contact.JobTitle,
                Department = contact.Department,
                OfficeLocation = contact.OfficeLocation,
                BusinessTelephoneNumber = contact.BusinessTelephoneNumber,
                MobileTelephoneNumber = contact.MobileTelephoneNumber,
                Domain = contact.Domain,
                RelationKinds = new List<string>(contact.RelationKinds ?? new List<string>()),
                SampleSubjects = new List<string>(contact.SampleSubjects ?? new List<string>()),
                MailCount = contact.MailCount,
                CalendarCount = contact.CalendarCount,
                FirstSeen = contact.FirstSeen,
                LastSeen = contact.LastSeen,
                RelationScore = contact.RelationScore,
                IsLikelySelf = contact.IsLikelySelf,
                IsKnown = contact.IsKnown,
                IsGroup = contact.IsGroup,
                MemberCount = contact.MemberCount,
                MemberSmtpAddresses = new List<string>(contact.MemberSmtpAddresses ?? new List<string>()),
                MemberGroupSmtpAddresses = new List<string>(contact.MemberGroupSmtpAddresses ?? new List<string>()),
                MemberOfGroupSmtpAddresses = new List<string>(contact.MemberOfGroupSmtpAddresses ?? new List<string>()),
            };
        }

        private static string ReadContactEmail(Outlook.ContactItem item, int emailSlot)
        {
            switch (emailSlot)
            {
                case 1: return ReadString(() => item.Email1Address);
                case 2: return ReadString(() => item.Email2Address);
                case 3: return ReadString(() => item.Email3Address);
                default: return "";
            }
        }

        private static string ReadContactEmailType(Outlook.ContactItem item, int emailSlot)
        {
            switch (emailSlot)
            {
                case 1: return ReadString(() => item.Email1AddressType);
                case 2: return ReadString(() => item.Email2AddressType);
                case 3: return ReadString(() => item.Email3AddressType);
                default: return "";
            }
        }

        private static string AddressListSource(Outlook.AddressList list)
        {
            try
            {
                return list.AddressListType == Outlook.OlAddressListType.olExchangeGlobalAddressList
                    ? "global_address_list"
                    : "outlook_address_list";
            }
            catch
            {
                return "address_list";
            }
        }

        private static int Clamp(int value, int min, int max, int fallback)
        {
            if (value <= 0) value = fallback;
            return Math.Max(min, Math.Min(max, value));
        }

        private static string ReadString(Func<string> read)
        {
            try { return read() ?? ""; } catch { return ""; }
        }

        private static string Prefer(string current, string candidate)
        {
            current = current ?? "";
            candidate = candidate ?? "";
            return string.IsNullOrWhiteSpace(current) ? candidate : current;
        }

        private static string Normalize(string value)
        {
            return (value ?? "").Trim().Trim('<', '>').ToLowerInvariant();
        }

        private static string EmailDomain(string email)
        {
            var at = (email ?? "").LastIndexOf('@');
            return at >= 0 && at < email.Length - 1 ? email.Substring(at + 1).ToLowerInvariant() : "";
        }

        private static string SmtpFromAddressEntry(Outlook.AddressEntry entry)
        {
            if (entry == null) return "";
            try
            {
                var user = entry.GetExchangeUser();
                if (user != null)
                {
                    var smtp = user.PrimarySmtpAddress ?? "";
                    Release(user);
                    if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                }
            }
            catch { }

            try
            {
                var distributionList = entry.GetExchangeDistributionList();
                if (distributionList != null)
                {
                    var smtp = distributionList.PrimarySmtpAddress ?? "";
                    Release(distributionList);
                    if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                }
            }
            catch { }

            return ReadString(() => entry.Address);
        }

        private static void Release(object obj)
        {
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
