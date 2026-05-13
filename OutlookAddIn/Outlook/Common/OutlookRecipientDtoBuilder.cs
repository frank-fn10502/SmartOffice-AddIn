using System.Collections.Generic;
using System.Runtime.InteropServices;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Common
{
    internal static class OutlookRecipientDtoBuilder
    {
        public static OutlookRecipientDto FromRecipient(Outlook.Recipient recipient, string kind)
        {
            var dto = new OutlookRecipientDto
            {
                RecipientKind = kind,
                Members = new List<OutlookRecipientDto>()
            };

            if (recipient == null)
                return dto;

            try { dto.DisplayName = recipient.Name ?? ""; } catch { dto.DisplayName = ""; }
            try { dto.RawAddress = recipient.Address ?? ""; } catch { dto.RawAddress = ""; }

            Outlook.AddressEntry addressEntry = null;
            try { addressEntry = recipient.AddressEntry; } catch { }
            if (addressEntry != null)
            {
                try
                {
                    try { dto.SmtpAddress = ResolveSmtpFromAddressEntry(addressEntry); } catch { }
                    try { dto.AddressType = addressEntry.AddressEntryUserType.ToString(); } catch { dto.AddressType = ""; }
                    try { dto.EntryUserType = addressEntry.AddressEntryUserType.ToString(); } catch { dto.EntryUserType = ""; }
                    try { dto.IsGroup = addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry; } catch { }
                    dto.IsResolved = true;
                }
                finally
                {
                    Release(addressEntry);
                }
            }

            if (string.IsNullOrEmpty(dto.SmtpAddress)) dto.SmtpAddress = "";
            if (string.IsNullOrEmpty(dto.AddressType)) dto.AddressType = "";
            if (string.IsNullOrEmpty(dto.EntryUserType)) dto.EntryUserType = "";
            return dto;
        }

        public static OutlookRecipientDto FromSender(Outlook.MailItem mail)
        {
            var dto = new OutlookRecipientDto
            {
                RecipientKind = "sender",
                Members = new List<OutlookRecipientDto>()
            };

            if (mail == null)
                return dto;

            try { dto.DisplayName = mail.SenderName ?? ""; } catch { dto.DisplayName = ""; }
            try { dto.RawAddress = mail.SenderEmailAddress ?? ""; } catch { dto.RawAddress = ""; }

            Outlook.AddressEntry addressEntry = null;
            try { addressEntry = mail.Sender; } catch { }
            if (addressEntry != null)
            {
                try
                {
                    try { dto.SmtpAddress = ResolveSmtpFromAddressEntry(addressEntry); } catch { }
                    try { dto.AddressType = addressEntry.AddressEntryUserType.ToString(); } catch { dto.AddressType = ""; }
                    try { dto.EntryUserType = addressEntry.AddressEntryUserType.ToString(); } catch { dto.EntryUserType = ""; }
                    try { dto.IsGroup = addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry; } catch { }
                    dto.IsResolved = true;
                }
                finally
                {
                    Release(addressEntry);
                }
            }

            if (string.IsNullOrEmpty(dto.SmtpAddress)) dto.SmtpAddress = "";
            if (string.IsNullOrEmpty(dto.AddressType)) dto.AddressType = "";
            if (string.IsNullOrEmpty(dto.EntryUserType)) dto.EntryUserType = "";
            return dto;
        }

        private static string ResolveSmtpFromAddressEntry(Outlook.AddressEntry addressEntry)
        {
            if (addressEntry == null)
                return "";

            try
            {
                if (addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                {
                    Outlook.ExchangeUser exchangeUser = null;
                    try
                    {
                        exchangeUser = addressEntry.GetExchangeUser();
                        if (exchangeUser != null)
                        {
                            var smtp = exchangeUser.PrimarySmtpAddress ?? "";
                            if (!string.IsNullOrEmpty(smtp))
                                return smtp;
                        }
                    }
                    finally
                    {
                        Release(exchangeUser);
                    }
                }
            }
            catch { }

            try { return addressEntry.Address ?? ""; } catch { return ""; }
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
