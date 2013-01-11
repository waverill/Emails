namespace Emails
{
    using System;
    using System.Globalization;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using System.Linq;
    using System.IO;
    using System.Text;
    using _MI_ = Microsoft.Office.Interop.Outlook.MailItem;
    using System.Windows.Forms;
    using System.Net;

    /// <summary>
    /// This class provides a simplifed interface to Outlook for the IAV e-mail project
    /// </summary>
    public class OutlookConnect
    {

        //public const string UNZIP_PATH = @"C:\Users\waverill\Documents\My Received Files";

        public OutlookConnect()
        {

        }
        private string make_final_body(string iava_number, string iava_name, string buildnum, List<string> audits)
        {
            string auditz = "";
            foreach (string a in audits)
            {
                auditz += a;
            }

            string body = "This email confirms delivery of vulnerability audits associated with the below referenced release.\r\n\r\n";
            body += "Notes\r\n- None\r\n\r\n";
            body += "IAV ID\r\n";
            body += "- " + iava_number;
            body += "\r\n\r\n";
            body += "IAV Title\r\n";
            body += "- " + iava_name;
            body += "\r\n\r\nAudit Revision\r\n- " + buildnum + "\r\n\r\n";
            body += "Audit Name (RTH ID)\r\n";
            body += auditz + "\r\n";
            body += "Release Date\r\n";
            body += "- " + DateTime.Now.ToString("MMMM dd, yyyy");
            body += "\r\n\r\nSigned,\r\nAudits Team\r\nEngineering Department\r\neEye Digital Security\r\n\r\n";
            body += "----------------------------------------------------------\r\n\r\n";
            return body;
        }

        private string get_iavas()
        {
            return null;
        }

        private string get_latest_build()
        {
            try
            {
                String builds_path = @"\\dev-builds\Store\RetinaAudits\Install\";
                var dir = new DirectoryInfo(builds_path);
                var build_dir = (from f in dir.GetDirectories()
                                 orderby f.LastWriteTime descending
                                 select f).First();
                builds_path += build_dir.ToString();
                builds_path += @"\Release";
                var dir2 = new DirectoryInfo(builds_path);
                FileInfo[] fi = dir2.GetFiles();
                Match m;
                foreach (FileInfo f in fi)
                {
                    m = Regex.Match(f.ToString(), @".*IAV_.*\.txt$", RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        String match1 = f.ToString();
                        Console.WriteLine(match1);
                        return match1.Substring(match1.IndexOf('_') + 1, 4);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return "ERROR";
        }

        private string get_build_num()
        {
            string url = "http://www.eeye.com/audits/latestversion.aspx";
            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(url);
            myRequest.Method = "GET";
            WebResponse myResponse = myRequest.GetResponse();
            StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8);
            string result = sr.ReadToEnd();
            sr.Close();
            myResponse.Close();
            Regex build = new Regex(@"([0-9]{4})");
            Match m = build.Match(result);
            return m.Value;
        }
        protected String Construct_Final(string html_body, string iava_name, string iava_number, List<string> audits, Boolean save)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
           // Regex exp = new Regex(@"Title:.*\r\n");
           // Regex exp2 = new Regex(@"IAVM Notice Number:.*");

          /*  _MI_ temp_message = null;
            temp_message = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            temp_message.HTMLBody = html_body;
            string plaintext_body = temp_message.Body;
            MatchCollection Matches = exp.Matches(plaintext_body);
            MatchCollection Matches2 = exp2.Matches(plaintext_body);*/

            _MI_ finalmessage = null;
            finalmessage = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
           // string iava_name = Matches[0].Value.Substring(Matches[0].Value.IndexOf(":") + 1).Trim();
           // string iava_number = Matches2[0].Value.Substring(Matches2[0].Value.IndexOf(":") + 1).Trim().Substring(0, 11);
            string subj = "Acknowledgement: Release of " + iava_number + " - " + iava_name;
            finalmessage.Body = make_final_body(iava_number, iava_name, this.get_latest_build(), audits) + html_body;
            finalmessage.Subject = subj;
            finalmessage.To = "iava@eeye.com";
           // MessageBox.Show(this.get_build_num());
            if (save)
            this.Save_to_Draftbox(finalmessage);

            return finalmessage.Body.ToString();
            
        }

        protected List<_MI_> Search_Sender(string subfolder_name, string query, bool only_unread)
        {
            return this._Search_Subfolder(subfolder_name, query, 0, only_unread);
        }

        protected List<_MI_> Search_Subject(string subfolder_name, string query, bool only_unread)
        {
            return this._Search_Subfolder(subfolder_name, query, 1, only_unread);
        }

        protected List<_MI_> Search_Bodies(string subfolder_name, string query, bool only_unread)
        {
            return this._Search_Subfolder(subfolder_name, query, 2, only_unread);
        }

        /// <summary>
        /// Saves an email object of type Microsoft.Office.Interop.Outlook.MailItem to Drafts Box
        /// </summary>
        /// <param name="new_message">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <returns>True/False on Success/Failure</returns>
        protected bool Save_to_Draftbox(_MI_ new_message)
        {
            Microsoft.Office.Interop.Outlook.Application app = null;
            Microsoft.Office.Interop.Outlook._NameSpace nameSpace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder drafts = null;
            try
            {
                app = new Microsoft.Office.Interop.Outlook.Application();
                nameSpace = app.GetNamespace("MAPI");
                nameSpace.Logon(null, null, false, false);

                drafts = nameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderDrafts);
                new_message.Save();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }

            return true;
        }

        /// <summary>
        /// Takes an Outlook email object and searches the Subject field
        /// </summary>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Subject</param>
        /// <returns>True/False on Found/Not found</returns>
        protected bool Search_Subject(_MI_ email_obj, string query)
        {
            if (email_obj.Subject.ToLower().Contains(query.ToLower()))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Takes an Outlook email object and searches the SenderName field
        /// </summary>
        /// <remarks>We need this because internal emails have messed up SenderEmail field</remarks>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Sender</param>
        /// <returns>True/False on Found/Not found</returns>
        protected bool Search_Sender_Name(_MI_ email_obj, string query)
        {
            if (email_obj.SenderName.ToLower().Contains(query.ToLower()))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Takes an Outlook email object and searches the Sender field
        /// </summary>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Sender</param>
        /// <returns>True/False on Found/Not found</returns>
        protected bool Search_Sender(_MI_ email_obj, string query)
        {
            if (email_obj.SenderEmailAddress.ToLower().Contains(query.ToLower()))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Takes an Outlook email object and searches the Attachment field
        /// </summary>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Attached file names</param>
        /// <returns>True/False on Found/Not found</returns>
        protected bool Search_Attachment_Name(_MI_ email_obj, string query)
        {
            System.Collections.IEnumerator enumer = null;
            enumer = email_obj.Attachments.GetEnumerator();
            while (enumer.MoveNext())
            {
                Microsoft.Office.Interop.Outlook.Attachment temp = null;
                temp = (Microsoft.Office.Interop.Outlook.Attachment)enumer.Current;
                if (temp.FileName.ToLower().Contains(query.ToLower()))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Takes an Outlook email object and searches the Body field
        /// </summary>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Body</param>
        /// <returns>True/False on Found/Not found</returns>
        protected bool Search_Body(_MI_ email_obj, string query)
        {
            if (email_obj.Body.ToLower().Contains(query.ToLower()))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Takes an Outlook email object saves the attachemnts
        /// </summary>
        /// <param name="email_obj">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <param name="query">String to search for in Subject</param>
        /// <returns>True/False on Success/Failure</returns>
   /*     protected string Save_Attachments(_MI_ email_obj, bool hasFinalNotices)
        {
            System.Collections.IEnumerator enumer = email_obj.Attachments.GetEnumerator();
            int count = 1;

            while (enumer.MoveNext())
            {

                if (email_obj.Attachments[count].FileName.Contains(".zip") || email_obj.Attachments[count].FileName.Contains(".ZIP"))
                {
                    email_obj.Attachments[count].SaveAsFile(UNZIP_PATH + email_obj.Attachments[count].FileName);
                    Console.WriteLine("Saving file as {0}", email_obj.Attachments[count].FileName);
                    Console.WriteLine("We got the IAVA Finals. Unzipping");
                    Console.WriteLine("Attempting to unzip {0}", UNZIP_PATH + email_obj.Attachments[count].FileName);
                    LocalStorage ls = new LocalStorage(UNZIP_PATH);
                    string daterange = ls.CreateDateRangeName();
                    ls.ExtractZipFile(UNZIP_PATH + email_obj.Attachments[count].FileName, daterange);
                    return UNZIP_PATH + daterange;
                }
                count++;
            }

            return null;
        }*/

        /// <summary>
        /// Search a specified Subfolder in Outlook
        /// </summary>
        /// <param name="subfolder_name">Name of Outlook Subfolder to search</param>
        /// <param name="query">Text to query</param>
        /// <param name="search_location">Location to search: {(0, Sender), (1, Subject), (2, Body)}</param>
        /// <param name="only_unread">True/False - Search only Unread Messages/Search All</param>
        /// <returns>A list of _MI_ objects representing matching emails</returns>
        private List<_MI_> _Search_Subfolder(string subfolder_name, string query, int search_location, bool only_unread)
        {
            List<_MI_> results = new List<_MI_>();
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook._NameSpace nameSpace = null;
                Microsoft.Office.Interop.Outlook.MAPIFolder inboxFolder = null;
                Microsoft.Office.Interop.Outlook.MAPIFolder subFolder = null;
                _MI_ item = null;
                nameSpace = app.GetNamespace("MAPI");
                nameSpace.Logon(null, null, false, false);
                inboxFolder = nameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                try
                {
                    subFolder = inboxFolder.Folders[subfolder_name];
                }
                catch (Exception not_found)
                {
                    Console.WriteLine(not_found.ToString());
                }

                for (int i = 1; i <= subFolder.Items.Count; i++)
                {
                    item = (_MI_)subFolder.Items[i];
                    query = query.Trim();
                    switch (search_location)
                    {
                        case 0:
                            Console.WriteLine("Sender: {0}", item.SenderEmailAddress);
                            Console.WriteLine("Other Sender: {0}", item.SenderName);
                            if (item.SenderEmailAddress.ToLower().Contains(query.ToLower()))
                            {
                                if ((only_unread == false) || (only_unread == true && item.UnRead == true))
                                {
                                    results.Add(item);
                                }
                            }

                            break;

                        case 1:
                            if (item.Subject.ToLower().Contains(query.ToLower()))
                            {
                                if ((only_unread == false) || (only_unread == true && item.UnRead == true))
                                {
                                    results.Add(item);
                                }
                            }

                            break;

                        case 2:
                            if (item.Body.ToLower().Contains(query.ToLower()))
                            {
                                if ((only_unread == false) || (only_unread == true && item.UnRead == true))
                                {
                                    results.Add(item);
                                }
                            }

                            break;

                        default:
                            throw new ArgumentException("Error in Query Location");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return results;
        }
    }
}