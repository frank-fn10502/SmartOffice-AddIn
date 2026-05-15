using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Contacts
{
    internal sealed partial class OutlookAddressBookReader
    {
        public List<AddressBookRootDto> ReadAddressBookRoots()
        {
            var roots = new List<AddressBookRootDto>();
            Outlook.AddressLists lists = null;
            try
            {
                lists = _application.Session.AddressLists;
                if (lists == null) return roots;

                for (var listIndex = 1; listIndex <= lists.Count; listIndex++)
                    AddAddressListRoot(roots, lists, listIndex);
            }
            finally
            {
                Release(lists);
            }

            return roots.OrderBy(root => root.Name).ToList();
        }

        public async Task<AddressBookListEntriesPageDto> ReadAddressListEntriesAsync(AddressBookListEntriesRequest request, string requestId)
        {
            if (request == null) request = new AddressBookListEntriesRequest();
            var offset = Math.Max(0, request.Offset);
            var pageSize = Math.Max(1, Math.Min(500, request.PageSize <= 0 ? 100 : request.PageSize));
            var contacts = new Dictionary<string, AddressBookContactDto>(StringComparer.OrdinalIgnoreCase);
            Outlook.AddressList list = null;
            Outlook.AddressEntries entries = null;
            try
            {
                list = FindAddressList(request);
                entries = list?.AddressEntries;
                if (list == null || entries == null)
                    return new AddressBookListEntriesPageDto { RequestId = requestId, Offset = offset, PageSize = pageSize };

                var total = entries.Count;
                var start = Math.Min(offset + 1, total + 1);
                var end = Math.Min(total, offset + pageSize);
                for (var entryIndex = start; entryIndex <= end; entryIndex++)
                {
                    Outlook.AddressEntry entry = null;
                    try
                    {
                        entry = entries[entryIndex];
                        await AddAddressEntryAsync(contacts, entry, list, 0, 0, new IntBudget(0));
                    }
                    catch { }
                    finally
                    {
                        Release(entry);
                    }
                    await YieldOutlookUiIfNeeded(entryIndex - offset);
                }

                return new AddressBookListEntriesPageDto
                {
                    RequestId = requestId,
                    AddressListId = ReadString(() => list.Name),
                    AddressListName = ReadString(() => list.Name),
                    Offset = offset,
                    PageSize = pageSize,
                    TotalCount = total,
                    HasMore = offset + contacts.Count < total,
                    Contacts = CloneContacts(contacts.Values),
                };
            }
            finally
            {
                Release(entries);
                Release(list);
            }
        }

        private void AddAddressListRoot(List<AddressBookRootDto> roots, Outlook.AddressLists lists, int listIndex)
        {
            Outlook.AddressList list = null;
            Outlook.AddressEntries entries = null;
            try
            {
                list = lists[listIndex];
                if (!IsSupportedAddressList(list)) return;

                entries = list.AddressEntries;
                roots.Add(new AddressBookRootDto
                {
                    Id = ReadString(() => list.Name),
                    Name = ReadString(() => list.Name),
                    AddressListType = ReadString(() => list.AddressListType.ToString()),
                    Source = AddressListSource(list),
                    EntryCount = entries?.Count ?? 0,
                    CanPageEntries = true,
                });
            }
            catch { }
            finally
            {
                Release(entries);
                Release(list);
            }
        }

        private Outlook.AddressList FindAddressList(AddressBookListEntriesRequest request)
        {
            Outlook.AddressLists lists = null;
            try
            {
                lists = _application.Session.AddressLists;
                if (lists == null) return null;

                var requestedId = (request.AddressListId ?? string.Empty).Trim();
                var requestedName = (request.AddressListName ?? string.Empty).Trim();
                for (var listIndex = 1; listIndex <= lists.Count; listIndex++)
                {
                    var list = TryMatchAddressList(lists, listIndex, requestedId, requestedName);
                    if (list != null) return list;
                }
            }
            finally
            {
                Release(lists);
            }

            return null;
        }

        private Outlook.AddressList TryMatchAddressList(Outlook.AddressLists lists, int listIndex, string requestedId, string requestedName)
        {
            Outlook.AddressList list = null;
            try
            {
                list = lists[listIndex];
                if (!IsSupportedAddressList(list)) return null;

                var name = ReadString(() => list.Name);
                if (string.Equals(name, requestedId, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, requestedName, StringComparison.OrdinalIgnoreCase))
                    return list;

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (list != null
                    && !string.Equals(ReadString(() => list.Name), requestedId, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(ReadString(() => list.Name), requestedName, StringComparison.OrdinalIgnoreCase))
                    Release(list);
            }
        }
    }
}
