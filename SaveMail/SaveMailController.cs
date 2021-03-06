﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace SaveMail
{
    // Controller class representing the entry point for the plugin
    public partial class SaveMailController
    {
        private void SaveMailController_Load(object sender, RibbonUIEventArgs e)
        {
            SaveMailLogger.LogAction("Plugin Loaded!");
        }

        // Main method with at event listener for the plugin button
        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            SaveMailLogger.LogAction("Plugin Activated!");

            List<object> selectedEmails = SaveMailModel.GetSelectedEmails();
            SaveMailLogger.LogAction("Selected " + selectedEmails.Count + " emails.");

            Folder inbox = (Folder) new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            string inboxFolderName = "Saved Mail";

            CreateSavedMailFolder(inbox, inboxFolderName);
            
            bool moveItems = false;            
            
            if (selectedEmails.Count != 0)
            {
                string currentFolder = ((MailItem)selectedEmails[0]).Parent.FolderPath.ToString();

                SaveMailLogger.LogAction("Selected from: " + currentFolder);

                Dictionary<object, object> savePath = SaveMailView.ShowBrowserDialog();

                if ((DialogResult)savePath["dialogResult"] == DialogResult.OK && currentFolder.Substring(currentFolder.Length - "Inbox".Length).Equals("Inbox") && SaveMailView.Question())
                {
                    moveItems = true;
                    SaveMailLogger.LogAction("Attempting to move selected out of the inbox...");
                }

                SaveMailView.Confirmation(SaveSelected(savePath, selectedEmails, inbox, inboxFolderName, moveItems));
            }
            else
            {
                SaveMailView.Confirmation("invalidSelection");
            }
        }

        // Create a folder in user inbox
        public static string CreateSavedMailFolder(Folder inbox, string inboxFolderName)
        {
            try {
                Folder checkFolder = (Folder)inbox.Folders[inboxFolderName];
                return "inboxFolderExists";
            }
            catch(System.Runtime.InteropServices.COMException)
            {
                inbox.Folders.Add(inboxFolderName, OlDefaultFolders.olFolderInbox);
                SaveMailLogger.LogAction("Inbox Folder Created!");
                return "inboxFolderCreated";
            }
        }

        // Invoke a sanity check for the path and save e-mails to drive
        public static string SaveSelected(Dictionary<object, object> savePath, List<object> emailItems, Folder inbox, String inboxFolderName, Boolean moveItems)
        {
            int savedNumber = 0;

            foreach (MailItem email in emailItems)
            {
                string pathCheckResult = SaveMailModel.PathCheck(savePath, email);

                if (!pathCheckResult.Equals("pathInvalid") && !pathCheckResult.Equals("saveCancelled"))
                {
                    if (email.ReceivedByName == null)
                    {
                        string emailDestination = SaveMailModel.GetEmailAddress(email, "outgoing");
                        email.SaveAs(savePath["selectedPath"] + "\\" + email.ReceivedTime.ToString("yyyy-MM-dd HHmm") + " " + emailDestination + " " + pathCheckResult + ".msg", OlSaveAsType.olMSG);
                    }
                    else
                    {
                        string emailSender = SaveMailModel.GetEmailAddress(email, "incoming");
                        email.SaveAs(savePath["selectedPath"] + "\\" + email.ReceivedTime.ToString("yyyy-MM-dd HHmm") + " " + emailSender + " " + pathCheckResult + ".msg", OlSaveAsType.olMSG);
                    }

                    if (moveItems)
                    {
                        email.Move(inbox.Folders[inboxFolderName]);
                        SaveMailLogger.LogAction("Moving item to " + inboxFolderName + " folder");
                    }

                    SaveMailLogger.LogAction("Saving item " + (savedNumber+1) + " with " + email.Size.ToString() + " bytes");
                    savedNumber++;
                }
                else
                {
                    return pathCheckResult;
                }
            }

            SaveMailLogger.LogAction("Saved to location: " + savePath["selectedPath"]);
            SaveMailLogger.LogAction("Total saved: " + savedNumber);
            return "saveSuccess";
        }
    }
}
