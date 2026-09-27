using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookIngest;

public class OutlookClient : IDisposable
{
    private readonly Application _outlookApp;
    private readonly NameSpace _namespace;
    private bool _disposed;

    public OutlookClient()
    {
        _outlookApp = new Application();
        _namespace = _outlookApp.GetNamespace("mapi");
    }

    /// <summary>
    /// Returns the store with the shortest StoreID from the Outlook profile.
    /// (Local mailbox stores typically have shorter identifiers than shared/secondary stores.)
    /// </summary>
    public Store GetFirstStore()
    {
        var stores = _namespace.Stores;
        var storeList = new List<Store>();
        foreach (Store store in stores)
        {
            storeList.Add(store);
        }

        if (storeList.Count == 0)
        {
            throw new InvalidOperationException("No stores found in Outlook profile.");
        }

        return storeList.OrderBy(s => s.StoreID.Length).First();
    }

    /// <summary>
    /// Gets all MailItem objects from the Inbox folder.
    /// </summary>
    public IEnumerable<MailItem> GetInboxItems(Store store)
    {
        var rootFolder = _namespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        var items = rootFolder.Items;

        var mailItems = new List<MailItem>();
        foreach (var item in items)
        {
            if (item is MailItem mail)
            {
                mailItems.Add(mail);
            }
        }

        // Release COM objects for non-MailItem entries
        Marshal.ReleaseComObject(items);
        Marshal.ReleaseComObject(rootFolder);

        return mailItems;
    }

    /// <summary>
    /// Gets all MailItem objects from the Sent Messages folder.
    /// </summary>
    public IEnumerable<MailItem> GetSentItems(Store store)
    {
        var rootFolder = _namespace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
        var items = rootFolder.Items;

        var mailItems = new List<MailItem>();
        foreach (var item in items)
        {
            if (item is MailItem mail)
            {
                mailItems.Add(mail);
            }
        }

        // Release COM objects for non-MailItem entries
        Marshal.ReleaseComObject(items);
        Marshal.ReleaseComObject(rootFolder);

        return mailItems;
    }

    public void Dispose()
    {
        if (_disposed) return;

        if (_namespace != null)
            Marshal.ReleaseComObject(_namespace);
        if (_outlookApp != null)
            Marshal.ReleaseComObject(_outlookApp);

        _disposed = true;
    }
}
