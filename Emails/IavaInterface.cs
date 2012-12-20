using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using _MI_ = Microsoft.Office.Interop.Outlook.MailItem;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

namespace Emails
{
    class IavaInterface : OutlookConnect
    {
        private string IAVA_ID;
        private string IAVA_Title;
        private List<string> auditIDs;
       
        public IavaInterface(string id, string title)
        {
           this.IAVA_ID = id;
           this.IAVA_Title = title;
           this.auditIDs = new List<string>(); 
        
        }

       /* public void addIava(string id, string title)
        {
            this.IAVA_ID[num] = id;
            this.IAVA_Title[num] = title;
            this.num++;
        }*/

        /// <summary>
        /// Takes an Outlook mailobject and parses out the IAVA id from the subject
        /// </summary>
        /// <param name="iava">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <returns>String containing IAVA ID</returns>
        public string Get_Iava_ID()
        {
            return this.IAVA_ID;
        }

        /// <summary>
        /// Takes an Outlook mailobject and parses out the IAVA name
        /// </summary>
        /// <param name="iava">Microsoft.Office.Interop.Outlook.MailItem object</param>
        /// <returns>String containing IAVA name</returns>
        public string Get_Iava_Name()
        {
            return this.IAVA_Title;
        }

        /// <summary>
        /// Returns the template used in IAVA Acknowledgement Emails
        /// </summary>
        /// <returns>String containing the IAVA Template</returns>
        public string iava_ack_template()
        {
            System.Text.StringBuilder iava_email = new System.Text.StringBuilder();
            iava_email.Append("This email confirms delivery of vulnerability audits associated with the below referenced release.\r\n\r\n");
            iava_email.Append("Notes\r\n- None\r\n\r\n");
            iava_email.Append("IAV ID\r\n- ");
            iava_email.Append(this.IAVA_ID);
            iava_email.Append("\r\n\r\nIAV Title\r\n- ");
            iava_email.Append(this.IAVA_Title);
            iava_email.Append("\r\n\r\nAudit Revision\r\n- TBD\r\n\r\n");
            iava_email.Append("Audit Name (RTH ID)\r\n- TBD\r\n\r\nRelease Date\r\n- TBD\r\n\r\n");
            iava_email.Append("Signed,\r\nAudits Team\r\nEngineering Department\r\neEye Digital Security\r\n\r\n");
            iava_email.Append("----------------------------------------------------------\r\n\r\n\r\n");
            return iava_email.ToString();
        }

        /// <summary>
        /// Takes list of IAVA precoords and constructs replies, saving them in the draftbox
        /// </summary>
        /// <param name="iavas">List of IAVAs of type Microsoft.Office.Interop.Outlook.MailItem</param>
        /// <returns>Number of IAVA acknowledements constructed</returns>
        public void Make_IAVA_Pre_ACK(string temp)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
           //_MI_ reply = ((Microsoft.Office.Interop.Outlook._MailItem)iava).Reply();
            Microsoft.Office.Interop.Outlook.MailItem reply = (Microsoft.Office.Interop.Outlook.MailItem) app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            reply.To = "iava@eeye.com";
           // Console.WriteLine(reply.To);
            reply.Subject = "Acknowledgement: Precoordination of " + this.IAVA_ID + " - " + this.IAVA_Title;
            reply.Body = temp;
            this.Save_to_Draftbox(reply);
        }


        /// <summary>
        /// We got the precoords in a box. acknowledge them
        /// </summary>
        /// <param name="foldername"></param>
        public void DoPreCoords(string foldername)
        {
            List<_MI_> temp = Search_Sender("TEST", "IAVM", true);
         //   foreach (_MI_ iava in temp)
        //        this.Make_IAVA_Pre_ACK(iava);
        }

     /*   public void DoFinalAcks(string precoord_folder, string this_week_iav_folder, string patt)
        {
            ///GET a ref to the zip file email
            List<_MI_> temp = Search_Subject(this_week_iav_folder, patt, true);
            string iav_path;
            Console.Out.WriteLine("[*]Found Final Notices");
            iav_path = Save_Attachments(temp[0], false);
            Console.WriteLine("iav path is {0}", iav_path);
            DirectoryInfo dir = new DirectoryInfo(iav_path);
            foreach(FileInfo f in dir.GetFiles("*.htm"))
            {
               Console.WriteLine(iav_path+f.Name);
               StreamReader sr = new StreamReader(iav_path + "\\" + f.Name);
               string html_content = sr.ReadToEnd();
               Construct_Final(html_content);
            }

        }*/
        public void DoFinalAcks2(string zip_path)
        {
            LocalStorage ls = new LocalStorage(zip_path);
            string daterange = ls.CreateDateRangeName();
            ls.ExtractZipFile(zip_path+"\\eEyeAuditUpdateIAV_2583-2584-en.zip", daterange);
            DirectoryInfo dir = new DirectoryInfo(zip_path+"\\"+daterange);
            foreach (FileInfo f in dir.GetFiles("*.txt"))
            {
                Console.WriteLine(zip_path+ f.Name);
                StreamReader sr = new StreamReader(zip_path+ "\\"+ daterange+"\\"+ f.Name);
                string html_content = sr.ReadToEnd();
                MessageBox.Show(html_content);
      //          Construct_Final(html_content,this.IAVA_ID, this.IAVA_Title);
            }
        }

        public void DoFinalAcks3(string content, string patchpath)
        {
           List<string> j = getAudits(patchpath);
           List<string> k = new List<string> ();
           foreach (string a in j)
           {
               string end = a.Substring(15);
               string beg = "- " + end;
               string audit_id = a.Substring(8, 5);
               beg += " (" + audit_id + ")\r\n";
              // MessageBox.Show(beg);
               k.Add(beg);
           }
           Construct_Final(content,this.IAVA_Title,this.IAVA_ID, k);
        }

        private List<string> getAudits(string patchpath)
        {
            List<string> audits = new List<string>();
            //int counter = 0;
           // string b = "@(\[" + this.IAVA_ID + "\]" + " " + this.IAVA_Title + ")";
           // MessageBox.Show(b);
            Regex yep = new Regex(@"(\[" + this.IAVA_ID + @"\] " + this.IAVA_Title + ")");
            Regex iav = new Regex(@"(==.*$)");
           /*   // Read the file and display it line by line.
            System.IO.StreamReader f = new System.IO.StreamReader(patchpath);
            while ((line = f.ReadLine()) != null)
            {
                Console.WriteLine(line);
                counter++;
            }
            
            f.Close();*/
            // Read in lines
            string[] lines = File.ReadAllLines(patchpath);
            Boolean we_there_yet = false;
            // Iterate through lines
            foreach (string l in lines)
            {                
                Match m2 = yep.Match(l);
                Match m3 = iav.Match(l);
                if (m3.Success && we_there_yet == true)
                {
                   // MessageBox.Show(l);
                    audits.Add(l);
                }
                else if (!m3.Success && we_there_yet == true)
                {
                  //  MessageBox.Show("Oh shit got 'em all!");
                    break;
                }
                if (m2.Success)
                {
                    we_there_yet = true;
                    
                }
            }
            return audits;
        }
    }
}
