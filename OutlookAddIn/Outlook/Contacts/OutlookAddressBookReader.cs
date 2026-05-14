using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Contacts
{
    internal sealed partial class OutlookAddressBookReader
    {
        private readonly Outlook.Application _application;
        private const int UiYieldInterval = 10;
        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const string PrEmailAddress = "http://schemas.microsoft.com/mapi/proptag/0x3003001E";

        public OutlookAddressBookReader(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public async Task<List<AddressBookContactDto>> ReadAddressBookAsync(AddressBookSyncRequest request, Action<List<AddressBookContactDto>> publishSnapshot = null)
        {
            request = request ?? new AddressBookSyncRequest();
            var maxContacts = ClampOrDefault(request.MaxContacts, 5000);
            var maxAddressEntriesPerList = ClampOrDefault(request.MaxAddressEntriesPerList, 2000);
            var maxGroupMembers = request.MaxGroupMembers < 0 ? 0 : Math.Min(request.MaxGroupMembers, 500);
            var maxGroupDepth = request.MaxGroupDepth < 0 ? 1 : Math.Min(request.MaxGroupDepth, 3);
            var groupMemberReadBudget = new IntBudget(Math.Min(maxContacts, 1000));
            var contacts = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);
            var publishThreshold = 50;
            var lastPublishedCount = new PublishCounter(0);

            if (request.IncludeOutlookContacts)
                await ReadDefaultContactsFolderAsync(contacts, maxContacts, publishSnapshot, publishThreshold, lastPublishedCount);

            if (request.IncludeAddressLists && contacts.Count < maxContacts)
                await ReadAddressListsAsync(contacts, maxContacts, maxAddressEntriesPerList, maxGroupMembers, maxGroupDepth, groupMemberReadBudget, publishSnapshot, publishThreshold, lastPublishedCount);

            var result = contacts.Values
                .OrderBy(item => item.DisplayName)
                .ThenBy(item => item.SmtpAddress)
                .Take(maxContacts)
                .ToList();
            publishSnapshot?.Invoke(CloneContacts(result));
            return result;
        }

        public async Task<List<AddressBookContactDto>> ReadAddressBookGroupMembersAsync(AddressBookGroupMembersRequest request)
        {
            request = request ?? new AddressBookGroupMembersRequest();
            var maxMembers = ClampOrDefault(request.MaxMembers, 5000);
            Outlook.ExchangeDistributionList distributionList = null;
            Outlook.AddressEntries members = null;
            var result = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);

            try
            {
                distributionList = FindExchangeDistributionList(request);
                if (distributionList == null) return new List<AddressBookContactDto>();

                members = distributionList.GetExchangeDistributionListMembers();
                if (members == null) return new List<AddressBookContactDto>();

                var limit = Math.Min(members.Count, maxMembers);
                for (var i = 1; i <= limit; i++)
                {
                    Outlook.AddressEntry member = null;
                    try
                    {
                        member = members[i];
                        var dto = await AddressEntryToContactAsync(member);
                        if (dto != null) Upsert(result, dto);
                    }
                    catch { }
                    finally
                    {
                        Release(member);
                    }
                    await YieldOutlookUiIfNeeded(i);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadAddressBookGroupMembers error: " + ex);
            }
            finally
            {
                Release(members);
                Release(distributionList);
            }

            return result.Values
                .OrderBy(item => item.DisplayName)
                .ThenBy(item => item.SmtpAddress)
                .ToList();
        }

        public async Task<AddressBookRelationLookupResponse> ReadAddressBookRelationLookupAsync(AddressBookRelationLookupRequest request)
        {
            request = request ?? new AddressBookRelationLookupRequest();
            request.Take = request.Take <= 0 ? 50 : Math.Min(request.Take, 500);
            var response = new AddressBookRelationLookupResponse
            {
                Query = RelationQuery(request),
                TargetKind = request.TargetKind ?? string.Empty,
                State = "not_found",
            };

            Outlook.AddressEntry entry = null;
            try
            {
                entry = ResolveRelationAddressEntry(request);
                if (entry == null)
                {
                    response.Matches = await FindRelationMatchesAsync(request);
                    response.State = response.Matches.Count > 1 ? "ambiguous" : "not_found";
                    if (response.Matches.Count == 1)
                    {
                        response.Target = response.Matches[0];
                        response.State = "found";
                    }
                    return response;
                }

                response.Target = await AddressEntryToContactAsync(entry, "relation_lookup");
                if (response.Target == null) return response;

                response.State = "found";
                response.IsGroup = response.Target.IsGroup;
                if (response.Target.IsGroup)
                    await FillGroupRelationAsync(response, entry, request.Take);
                else
                    await FillUserRelationAsync(response, entry, request.Take);

                return response;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadAddressBookRelationLookup error: " + ex);
                response.Message = ex.Message;
                return response;
            }
            finally
            {
                Release(entry);
            }
        }

        private async Task ReadDefaultContactsFolderAsync(
            Dictionary<string, AddressBookContactDto> contacts,
            int maxContacts,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            PublishCounter lastPublishedCount)
        {
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;

            try
            {
                folder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                items = folder?.Items;
                if (items == null) return;

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
                        PublishIfNeeded(contacts, publishSnapshot, publishThreshold, lastPublishedCount);
                    }
                    catch { }
                    finally
                    {
                        Release(contact);
                        if (!ReferenceEquals(item, contact)) Release(item);
                    }
                    await YieldOutlookUiIfNeeded(i);
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

        private async Task ReadAddressListsAsync(
            Dictionary<string, AddressBookContactDto> contacts,
            int maxContacts,
            int maxAddressEntriesPerList,
            int maxGroupMembers,
            int maxGroupDepth,
            IntBudget groupMemberReadBudget,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            PublishCounter lastPublishedCount)
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
                        for (var entryIndex = 1; entryIndex <= entryLimit && contacts.Count < maxContacts; entryIndex++)
                        {
                            Outlook.AddressEntry entry = null;
                            try
                            {
                                entry = entries[entryIndex];
                                await AddAddressEntryAsync(contacts, entry, list, maxGroupMembers, maxGroupDepth, groupMemberReadBudget);
                                PublishIfNeeded(contacts, publishSnapshot, publishThreshold, lastPublishedCount);
                            }
                            catch { }
                            finally
                            {
                                Release(entry);
                            }
                            await YieldOutlookUiIfNeeded(entryIndex);
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

        private static async Task AddAddressEntryAsync(
            Dictionary<string, AddressBookContactDto> contacts,
            Outlook.AddressEntry entry,
            Outlook.AddressList list,
            int maxGroupMembers,
            int maxGroupDepth,
            IntBudget groupMemberReadBudget)
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

            var userType = ReadAddressEntryUserType(entry);
            if (!LooksLikeSmtpAddress(dto.SmtpAddress))
                dto.SmtpAddress = "";
            dto.SmtpAddress = Prefer(SmtpFromAddressEntryProperties(entry), dto.SmtpAddress);
            dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
            dto.IsGroup = IsDistributionListEntry(userType);

            Outlook.ExchangeUser exchangeUser = null;
            Outlook.ExchangeDistributionList distributionList = null;
            try
            {
                if (!LooksLikeSmtpAddress(dto.SmtpAddress) && IsExchangeUserEntry(userType))
                {
                    exchangeUser = entry.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => exchangeUser.Name));
                        dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => exchangeUser.PrimarySmtpAddress));
                        dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                    }
                }

                if (maxGroupMembers > 0 && IsDistributionListEntry(userType))
                {
                    distributionList = entry.GetExchangeDistributionList();
                    if (distributionList != null)
                    {
                        dto.IsGroup = true;
                        dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => distributionList.Name));
                        dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => distributionList.PrimarySmtpAddress));
                        dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                        await ReadDistributionListMembersAsync(distributionList, dto, maxGroupMembers, maxGroupDepth, 0, new HashSet<string>(StringComparer.OrdinalIgnoreCase), groupMemberReadBudget);
                    }
                }
            }
            catch { }
            finally
            {
                Release(distributionList);
                Release(exchangeUser);
            }

            if (!string.IsNullOrWhiteSpace(dto.SmtpAddress) || !string.IsNullOrWhiteSpace(dto.DisplayName))
                Upsert(contacts, dto);
        }

        private async Task<AddressBookContactDto> AddressEntryToContactAsync(Outlook.AddressEntry entry, string source = "group_member")
        {
            if (entry == null) return null;
            var userType = ReadAddressEntryUserType(entry);
            var smtp = SmtpFromAddressEntryProperties(entry);
            var dto = new AddressBookContactDto
            {
                Id = ReadString(() => entry.ID),
                DisplayName = ReadString(() => entry.Name),
                SmtpAddress = smtp,
                RawAddress = string.IsNullOrWhiteSpace(smtp) ? ReadString(() => entry.Address) : smtp,
                AddressType = ReadString(() => entry.Type),
                EntryUserType = ReadString(() => entry.AddressEntryUserType.ToString()),
                Source = source,
                IsGroup = IsDistributionListEntry(userType),
            };

            Outlook.ExchangeUser exchangeUser = null;
            Outlook.ExchangeDistributionList distributionList = null;
            try
            {
                if (!LooksLikeSmtpAddress(dto.SmtpAddress) && IsExchangeUserEntry(userType))
                {
                    exchangeUser = entry.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => exchangeUser.Name));
                        dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => exchangeUser.PrimarySmtpAddress));
                        dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                    }
                }

                if (IsDistributionListEntry(userType))
                {
                    distributionList = entry.GetExchangeDistributionList();
                    if (distributionList != null)
                    {
                        dto.DisplayName = Prefer(dto.DisplayName, ReadString(() => distributionList.Name));
                        dto.SmtpAddress = Prefer(dto.SmtpAddress, ReadString(() => distributionList.PrimarySmtpAddress));
                        dto.RawAddress = Prefer(dto.RawAddress, dto.SmtpAddress);
                        dto.MemberCount = ReadMemberCount(distributionList);
                    }
                }
            }
            catch { }
            finally
            {
                Release(distributionList);
                Release(exchangeUser);
            }

            await Task.CompletedTask;
            return string.IsNullOrWhiteSpace(dto.SmtpAddress) && string.IsNullOrWhiteSpace(dto.DisplayName) ? null : dto;
        }

        private Outlook.AddressEntry ResolveRelationAddressEntry(AddressBookRelationLookupRequest request)
        {
            var query = RelationQuery(request);
            if (string.IsNullOrWhiteSpace(query)) return null;

            Outlook.Recipient recipient = null;
            try
            {
                recipient = _application.Session.CreateRecipient(query);
                if (recipient == null) return null;
                if (!recipient.Resolve() || !recipient.Resolved) return null;
                return recipient.AddressEntry;
            }
            catch
            {
                return null;
            }
            finally
            {
                Release(recipient);
            }
        }

        private async Task FillGroupRelationAsync(AddressBookRelationLookupResponse response, Outlook.AddressEntry groupEntry, int take)
        {
            Outlook.ExchangeDistributionList distributionList = null;
            try
            {
                distributionList = groupEntry.GetExchangeDistributionList();
                if (distributionList == null) return;

                response.Members = await ReadDistributionListDirectMembersAsync(distributionList, take);
                response.MemberGroups = response.Members
                    .Where(member => member.IsGroup)
                    .ToList();
                response.Target.MemberCount = Math.Max(response.Target.MemberCount, ReadMemberCount(distributionList));
                response.Target.MemberSmtpAddresses = response.Members
                    .Select(member => member.SmtpAddress)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(500)
                    .ToList();
                response.Target.MemberGroupSmtpAddresses = response.MemberGroups
                    .Select(member => member.SmtpAddress)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(500)
                    .ToList();
            }
            catch { }
            finally
            {
                Release(distributionList);
            }
        }

        private async Task FillUserRelationAsync(AddressBookRelationLookupResponse response, Outlook.AddressEntry userEntry, int take)
        {
            Outlook.ExchangeUser exchangeUser = null;
            Outlook.AddressEntries groups = null;
            try
            {
                exchangeUser = userEntry.GetExchangeUser();
                groups = exchangeUser?.GetMemberOfList();
                if (groups == null) return;

                var limit = Math.Min(groups.Count, Math.Max(1, Math.Min(500, take)));
                for (var i = 1; i <= limit; i++)
                {
                    Outlook.AddressEntry groupEntry = null;
                    try
                    {
                        groupEntry = groups[i];
                        var group = await AddressEntryToContactAsync(groupEntry, "member_of_group");
                        if (group != null) response.MemberOfGroups.Add(group);
                    }
                    catch { }
                    finally
                    {
                        Release(groupEntry);
                    }
                    await YieldOutlookUiIfNeeded(i);
                }

                response.Target.MemberOfGroupSmtpAddresses = response.MemberOfGroups
                    .Select(group => group.SmtpAddress)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(500)
                    .ToList();
                response.ContainingGroups = response.MemberOfGroups.ToList();
            }
            catch { }
            finally
            {
                Release(groups);
                Release(exchangeUser);
            }
        }

        private async Task<List<AddressBookContactDto>> ReadDistributionListDirectMembersAsync(Outlook.ExchangeDistributionList distributionList, int take)
        {
            var members = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);
            Outlook.AddressEntries entries = null;
            try
            {
                entries = distributionList.GetExchangeDistributionListMembers();
                if (entries == null) return new List<AddressBookContactDto>();
                var limit = Math.Min(entries.Count, Math.Max(1, Math.Min(500, take)));
                for (var i = 1; i <= limit; i++)
                {
                    Outlook.AddressEntry memberEntry = null;
                    try
                    {
                        memberEntry = entries[i];
                        var member = await AddressEntryToContactAsync(memberEntry, "group_member");
                        if (member != null) Upsert(members, member);
                    }
                    catch { }
                    finally
                    {
                        Release(memberEntry);
                    }
                    await YieldOutlookUiIfNeeded(i);
                }
            }
            finally
            {
                Release(entries);
            }

            return members.Values
                .OrderBy(member => member.DisplayName)
                .ThenBy(member => member.SmtpAddress)
                .ToList();
        }

        private async Task<List<AddressBookContactDto>> FindRelationMatchesAsync(AddressBookRelationLookupRequest request)
        {
            var query = Normalize(RelationQuery(request));
            if (string.IsNullOrWhiteSpace(query)) return new List<AddressBookContactDto>();

            var matches = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);
            Outlook.AddressLists lists = null;
            try
            {
                lists = _application.Session.AddressLists;
                if (lists == null) return new List<AddressBookContactDto>();
                var limit = Math.Max(1, Math.Min(25, request.Take <= 0 ? 10 : request.Take));
                var inspected = 0;
                for (var listIndex = 1; listIndex <= lists.Count && matches.Count < limit && inspected < 1000; listIndex++)
                {
                    Outlook.AddressList list = null;
                    Outlook.AddressEntries entries = null;
                    try
                    {
                        list = lists[listIndex];
                        if (!IsSupportedAddressList(list)) continue;
                        entries = list.AddressEntries;
                        if (entries == null) continue;
                        var entryLimit = Math.Min(entries.Count, 250);
                        for (var entryIndex = 1; entryIndex <= entryLimit && matches.Count < limit && inspected < 1000; entryIndex++)
                        {
                            inspected++;
                            Outlook.AddressEntry entry = null;
                            try
                            {
                                entry = entries[entryIndex];
                                if (!AddressEntryMatchesRelation(entry, query, request)) continue;
                                var dto = await AddressEntryToContactAsync(entry, AddressListSource(list));
                                if (dto != null) Upsert(matches, dto);
                            }
                            catch { }
                            finally
                            {
                                Release(entry);
                            }
                            await YieldOutlookUiIfNeeded(entryIndex);
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
            finally
            {
                Release(lists);
            }

            return matches.Values
                .OrderBy(match => match.DisplayName)
                .ThenBy(match => match.SmtpAddress)
                .ToList();
        }

        private Outlook.ExchangeDistributionList FindExchangeDistributionList(AddressBookGroupMembersRequest request)
        {
            Outlook.AddressLists lists = null;
            try
            {
                lists = _application.Session.AddressLists;
                if (lists == null) return null;

                for (var listIndex = 1; listIndex <= lists.Count; listIndex++)
                {
                    Outlook.AddressList list = null;
                    Outlook.AddressEntries entries = null;
                    try
                    {
                        list = lists[listIndex];
                        if (!IsSupportedAddressList(list)) continue;

                        entries = list.AddressEntries;
                        if (entries == null) continue;
                        for (var entryIndex = 1; entryIndex <= entries.Count; entryIndex++)
                        {
                            Outlook.AddressEntry entry = null;
                            try
                            {
                                entry = entries[entryIndex];
                                if (!AddressEntryMatchesGroup(entry, request)) continue;
                                var distributionList = entry.GetExchangeDistributionList();
                                if (distributionList != null) return distributionList;
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
            finally
            {
                Release(lists);
            }

            return null;
        }

        private static async Task ReadDistributionListMembersAsync(
            Outlook.ExchangeDistributionList distributionList,
            AddressBookContactDto dto,
            int maxGroupMembers,
            int maxGroupDepth,
            int depth,
            HashSet<string> visitedGroups,
            IntBudget groupMemberReadBudget)
        {
            if (distributionList == null || maxGroupMembers <= 0 || depth > maxGroupDepth || groupMemberReadBudget.Value <= 0) return;
            var groupKey = ReadString(() => distributionList.PrimarySmtpAddress);
            if (!string.IsNullOrWhiteSpace(groupKey) && !visitedGroups.Add(groupKey)) return;

            Outlook.AddressEntries members = null;
            try
            {
                members = distributionList.GetExchangeDistributionListMembers();
                if (members == null) return;
                var limit = Math.Min(Math.Min(members.Count, maxGroupMembers - dto.MemberSmtpAddresses.Count), groupMemberReadBudget.Value);
                for (var i = 1; i <= limit; i++)
                {
                    groupMemberReadBudget.Value--;
                    Outlook.AddressEntry member = null;
                    Outlook.ExchangeDistributionList nested = null;
                    try
                    {
                        member = members[i];
                        var smtp = SmtpFromAddressEntryProperties(member);
                        if (string.IsNullOrWhiteSpace(smtp)) smtp = ReadString(() => member.Address);
                        if (!string.IsNullOrWhiteSpace(smtp) && !dto.MemberSmtpAddresses.Contains(smtp, StringComparer.OrdinalIgnoreCase))
                            dto.MemberSmtpAddresses.Add(smtp);

                        if (depth < maxGroupDepth && IsDistributionListEntry(ReadAddressEntryUserType(member)))
                        {
                            nested = member.GetExchangeDistributionList();
                            if (nested != null && !string.IsNullOrWhiteSpace(smtp))
                            {
                                dto.MemberGroupSmtpAddresses.Add(smtp);
                                await ReadDistributionListMembersAsync(nested, dto, maxGroupMembers, maxGroupDepth, depth + 1, visitedGroups, groupMemberReadBudget);
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        Release(nested);
                        Release(member);
                    }
                    await YieldOutlookUiIfNeeded(i);
                }
                dto.MemberCount = Math.Max(dto.MemberCount, members.Count);
            }
            catch { }
            finally
            {
                Release(members);
            }
        }

        private static Task YieldOutlookUiIfNeeded(int processedCount)
        {
            return processedCount % UiYieldInterval == 0 ? Task.Delay(1) : Task.CompletedTask;
        }

        private sealed class PublishCounter
        {
            public PublishCounter(int value)
            {
                Value = value;
            }

            public int Value { get; set; }
        }

        private sealed class IntBudget
        {
            public IntBudget(int value)
            {
                Value = value;
            }

            public int Value { get; set; }
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
            dto.MemberOfGroupSmtpAddresses = dto.MemberOfGroupSmtpAddresses.Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();

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
            current.MemberOfGroupSmtpAddresses = current.MemberOfGroupSmtpAddresses.Concat(dto.MemberOfGroupSmtpAddresses).Distinct(StringComparer.OrdinalIgnoreCase).Take(50).ToList();
            if (!current.Sources.Contains(dto.Source)) current.Sources.Add(dto.Source);
        }

        private static void PublishIfNeeded(
            Dictionary<string, AddressBookContactDto> contacts,
            Action<List<AddressBookContactDto>> publishSnapshot,
            int publishThreshold,
            PublishCounter lastPublishedCount)
        {
            if (publishSnapshot == null) return;
            if (contacts.Count - lastPublishedCount.Value < publishThreshold) return;
            lastPublishedCount.Value = contacts.Count;
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
                IsRelatedToSelf = contact.IsRelatedToSelf,
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

        private static int ClampOrDefault(int value, int max)
        {
            return value <= 0 ? max : Math.Min(value, max);
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

        private static bool LooksLikeSmtpAddress(string value)
        {
            var address = value ?? "";
            var at = address.IndexOf('@');
            return at > 0 && at < address.Length - 1;
        }

        private static Outlook.OlAddressEntryUserType? ReadAddressEntryUserType(Outlook.AddressEntry entry)
        {
            try { return entry.AddressEntryUserType; } catch { return null; }
        }

        private static bool IsExchangeUserEntry(Outlook.OlAddressEntryUserType? userType)
        {
            return userType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                || userType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry;
        }

        private static bool IsDistributionListEntry(Outlook.OlAddressEntryUserType? userType)
        {
            return userType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry
                || userType == Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry;
        }

        private static bool AddressEntryMatchesGroup(Outlook.AddressEntry entry, AddressBookGroupMembersRequest request)
        {
            if (entry == null || !IsDistributionListEntry(ReadAddressEntryUserType(entry))) return false;

            var requestedId = (request.GroupId ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(requestedId)
                && string.Equals(ReadString(() => entry.ID), requestedId, StringComparison.OrdinalIgnoreCase))
                return true;

            var requestedSmtp = Normalize(request.GroupSmtpAddress);
            if (string.IsNullOrWhiteSpace(requestedSmtp)) return false;
            var smtp = Normalize(SmtpFromAddressEntryProperties(entry));
            if (!string.IsNullOrWhiteSpace(smtp) && string.Equals(smtp, requestedSmtp, StringComparison.OrdinalIgnoreCase))
                return true;

            return string.Equals(Normalize(ReadString(() => entry.Address)), requestedSmtp, StringComparison.OrdinalIgnoreCase);
        }

        private static bool AddressEntryMatchesRelation(Outlook.AddressEntry entry, string query, AddressBookRelationLookupRequest request)
        {
            if (entry == null || string.IsNullOrWhiteSpace(query)) return false;
            var groupOnly = string.Equals(request.TargetKind, "group", StringComparison.OrdinalIgnoreCase)
                || !string.IsNullOrWhiteSpace(request.GroupSmtpAddress)
                || !string.IsNullOrWhiteSpace(request.GroupId);
            if (groupOnly && !IsDistributionListEntry(ReadAddressEntryUserType(entry))) return false;

            var id = Normalize(ReadString(() => entry.ID));
            var name = Normalize(ReadString(() => entry.Name));
            var address = Normalize(ReadString(() => entry.Address));
            var smtp = Normalize(SmtpFromAddressEntryProperties(entry));
            return id == query
                || smtp == query
                || address == query
                || name == query
                || (!LooksLikeSmtpAddress(query) && name.Contains(query));
        }

        private static string RelationQuery(AddressBookRelationLookupRequest request)
        {
            return Prefer(
                Prefer(
                    Prefer(
                        Prefer(
                            Prefer(request.GroupSmtpAddress, request.SmtpAddress),
                            request.Email),
                        request.Id),
                    request.GroupId),
                Prefer(request.Query, request.DisplayName));
        }

        private static int ReadMemberCount(Outlook.ExchangeDistributionList distributionList)
        {
            Outlook.AddressEntries members = null;
            try
            {
                members = distributionList.GetExchangeDistributionListMembers();
                return members?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                Release(members);
            }
        }

        private static string SmtpFromAddressEntryProperties(Outlook.AddressEntry entry)
        {
            if (entry == null) return "";
            var smtp = ReadPropertyString(entry, PrSmtpAddress);
            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;

            smtp = ReadPropertyString(entry, PrEmailAddress);
            if (!string.IsNullOrWhiteSpace(smtp) && smtp.Contains("@")) return smtp;

            var address = ReadString(() => entry.Address);
            return address.Contains("@") ? address : "";
        }

        private static string ReadPropertyString(Outlook.AddressEntry entry, string schemaName)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = entry.PropertyAccessor;
                return accessor?.GetProperty(schemaName) as string ?? "";
            }
            catch
            {
                return "";
            }
            finally
            {
                Release(accessor);
            }
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
