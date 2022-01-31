using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTOAlert
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // get a reference to the inbox in Outlook
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            // add and event handler so that when an item is added
            // to the inbox we can react to that if necessary
            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        /// <summary>
        /// Something has been added to the inbox
        /// </summary>
        /// <param name="Item"></param> the item that was added
        void items_ItemAdd(object Item)
        {
            try
            {
                //get the item that was added
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                if (Item != null)
                {
                    // ToSearchStrings is set by the ribbon (Defaults for now to Abbott ARR V-Team)
                    var SearchStrings = Properties.Settings.Default.ToSearchStrings;
                    // see if the To field contains one of the search strings set in the ribbon
                    foreach (string toAddr in SearchStrings)
                    {
                        // if email message and the To: line as one of the addresses of concern
                        if (mail.MessageClass == "IPM.Note" && mail.To.ToLower().Contains(toAddr.ToLower()))
                        {
                            // create a sound
                            System.Media.SystemSounds.Exclamation.Play();

                            // Bring the mail item to the front!
                            mail.GetInspector.Activate(); // note if there is a dialog up, this will throw an exception!

                            // no need to test any more, it met one of the criteria
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            { //todo surface exception? for example:
              // if dialog is up it this will be hit if we try and activate the mail item
            };
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        #endregion
    }
}
