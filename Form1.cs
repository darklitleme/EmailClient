using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Runtime.Serialization.Formatters.Binary;


namespace email
{

    public partial class Form1 : Form
    {

        google goog = new google();
        System.Threading.Thread load;

        private static void loadingWindow()
        {
            try
            {
                Loading wind = new Loading();
                wind.ShowDialog();
            }
            catch { }
        }

        public Form1()
        {
            InitializeComponent();

            webBrowser1.Url = new Uri("file:\\\\\\" + Directory.GetCurrentDirectory() + "\\staticG.gif");

            //pre load saved messages
            loadInbox();


        }


        /// <summary>
        /// updates the inbox listbox
        /// </summary>
        private void loadInbox()
        {
            load = new System.Threading.Thread(loadingWindow);
            load.Start();

            try
            {
                goog.updateInbox();

                listBox1.Items.Clear();
                listBox1.Sorted = false;

                List<userMessages> list = new List<userMessages>(goog.allMessages.Values);
                list.OrderBy(userMessages => userMessages.DATE);
                listBox1.DataSource = list;

                listBox1.DisplayMember = "TITLE";

            }
            catch (Exception Exc)
            { Debug.WriteLine("1" + Exc); }

            try
            {

                load.Abort();
            }
            catch { }

        }

        //load inbox
        private void button1_Click(object sender, EventArgs e)
        {
            loadInbox();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                userMessages meg = (userMessages)listBox1.SelectedItem;
                //goog.allMessages[meg.ID] = new userMessages( goog.GetMessage(meg.ID) , "INBOX");
                webBrowser1.DocumentText = goog.allMessages[meg.ID].CONTENTS;

            }
            catch (Exception eee)
            { Debug.WriteLine("2 " + eee); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            load = new System.Threading.Thread(loadingWindow);
            load.Start();
            try
            {
                listBox1.Items.Clear();
                goog.loadSentMessages();
                foreach (KeyValuePair<string, userMessages> mess in goog.allMessages)
                {
                    if (mess.Value.FOLDER == "SENT")
                    {
                        listBox1.Items.Add(mess.Value);
                    }
                }
                
                listBox1.DisplayMember = "TITLE";

                
            }
            catch (Exception Exc)
            { Debug.WriteLine("1" + Exc); }

            load.Abort();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void txtbSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Equals("{ENTER}"))
            {
                listBox1.Items.Clear();

                foreach (userMessages mess in goog.allMessages.Values)
                {
                    if (mess.TITLE.Contains(txtbSearch.Text))
                    {
                        listBox1.Items.Add(mess);
                    }
                }
            }
        }

        private void txtbSearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtbSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.Equals("{ENTER}"))
            {
                listBox1.Items.Clear();

                foreach (userMessages mess in goog.allMessages.Values)
                {
                    if (mess.TITLE.Contains(txtbSearch.Text))
                    {
                        listBox1.Items.Add(mess);
                    }
                }
            }
        }
    }



    [Serializable]
    class google
    {
        
        public Dictionary<string, userMessages> allMessages = new Dictionary<string, userMessages>(); //user messages

        public void updateInbox()
        {
            try
            {
               loadMessagesFromFile();
                

                List<Google.Apis.Gmail.v1.Data.Message> responce = ListMessages("INBOX");
                int max = 0;
                foreach (Google.Apis.Gmail.v1.Data.Message emil in responce)
                {
                    try
                    {
                        if (allMessages.ContainsKey(emil.Id.ToString()) == false)
                        {
                            allMessages.Add(emil.Id.ToString(), new userMessages(GetMessage(emil.Id), "INBOX"));
                            allMessages.Add(emil.Id.ToString(), new userMessages(emil, "INBOX"));
                            Debug.Write("\n hi :) \n");
                            max++;
                        }
                    }
                    catch (Exception exc) { Debug.WriteLine("3 " + exc); }

                    //for testing, to save time//
                    if (max > 50)
                        break;
                    ///
                }
                saveMessagesToFile();
            }
            catch (Exception exc) { Debug.WriteLine("4 " + exc); }



        }





        //loads user inbox from web
        public void loadMessages()
        {
            List<Google.Apis.Gmail.v1.Data.Message> responce = ListMessages("INBOX");
            foreach (Google.Apis.Gmail.v1.Data.Message emil in responce)
            {

                try { allMessages.Add(emil.Id.ToString(), new userMessages(GetMessage(emil.Id),"INBOX") )  ; }
                catch(Exception exc) { Debug.WriteLine("3 " + exc); }


            }
        }
        //load send messages from web
        public void loadSentMessages()
        {
            List<Google.Apis.Gmail.v1.Data.Message> responce = ListMessages("SENT");
            foreach (Google.Apis.Gmail.v1.Data.Message emil in responce)
            {

                try { allMessages.Add(emil.Id.ToString(), new userMessages(GetMessage(emil.Id), "SENT")); }
                catch (Exception exc) { Debug.WriteLine("3 " + exc); }


            }
        }

        //
        public void saveMessagesToFile()
        {
            try
            {
                using (Stream stream = File.Open("data.bin", FileMode.Create))
                {
                    BinaryFormatter bin = new BinaryFormatter();
                    bin.Serialize(stream, allMessages);
                }
            }
            catch (IOException)
            {
            }
        }
        //
        public void loadMessagesFromFile()
        {
            try
            {
                using (Stream streamin = File.Open("data.bin", FileMode.Open))
                {
                    BinaryFormatter bin = new BinaryFormatter();
                    allMessages = (Dictionary<string, userMessages>)bin.Deserialize(streamin);
                }
            }
            catch (IOException)
            {
            }

        }

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Alex's Message Manager";
        
        UserCredential credential;

        public void start()
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }
        
        }





        /// <summary>
        /// List all Messages of the user's mailbox matching the query.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="query">String used to filter Messages returned.</param>
        public List<Google.Apis.Gmail.v1.Data.Message> ListMessages(string box)
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            List<Google.Apis.Gmail.v1.Data.Message> result = new List<Google.Apis.Gmail.v1.Data.Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
            request.LabelIds = box;
            //request.Q = query;
            long max = 200;
            request.MaxResults = max;

            do
            {
                try
                {
                    ListMessagesResponse response = request.Execute();
                    result.AddRange(response.Messages);
                    request.PageToken = response.NextPageToken;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(" 4 An error occurred: " + e.Message);
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return result;
        }

        // ...




        /// <summary>
        /// Retrieve a Message by ID.
        /// </summary>
        /// can be used to indicate the authenticated user.</param>
        /// <param name="messageId">ID of Message to retrieve.</param>
        public Google.Apis.Gmail.v1.Data.Message GetMessage(String messageId)
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            try
            {
                return service.Users.Messages.Get("me", messageId).Execute();
            }
            catch (Exception e)
            {
                Debug.WriteLine("5 An error occurred: " + e.Message);
            }

            return null;
        }


    }

    [Serializable]
    class userMessages
     {
        public string TO { get; set; }
        public string FROM { get; set; }
        public DateTime DATE { get; set; }
        public string CONTENTS { get; set; }
        public string TITLE { get; set; }
        public string SUBJECT { get; set; }
        public string ID { get; set; }
        public string SENDER { get; set; }
        public string REPLYTO { get; set; }
        public string FOLDER { get; set; }

        public userMessages(Google.Apis.Gmail.v1.Data.Message mess , string folder)
        {
            ID = mess.Id;

            FOLDER = folder;


            if (mess.Payload.Parts == null && mess.Payload.Body != null)
            {
                try
                {

                    CONTENTS = mess.Payload.Body.Data;
                    String codedBody = CONTENTS.Replace("-", "+");
                    codedBody = codedBody.Replace("_", "/");
                    byte[] data = Convert.FromBase64String(CONTENTS);
                    CONTENTS = Encoding.UTF8.GetString(data);
                }
                catch
                {
                    CONTENTS = getNestedParts(mess.Payload.Parts, "");
                }
            }
            else 
            {
                try
                {
                    CONTENTS = getNestedParts(mess.Payload.Parts, "");
                }
                catch
                {
                    CONTENTS = mess.Payload.Body.ToString();
                }
            }

            //headders
            foreach (MessagePartHeader prt in mess.Payload.Headers)
            {
                if (prt.Name == "Date")
                {
                    DATE = DateTime.Parse(prt.Value);
                }
                else if (prt.Name == "From")
                {
                    FROM = prt.Value;
                }
                else if (prt.Name == "Sender")
                {
                    SENDER = prt.Value;
                }
                else if (prt.Name == "Reply-To")
                {
                    REPLYTO = prt.Value;
                }
                else if (prt.Name == "To")
                {
                    TO = prt.Value;
                }
                else if (prt.Name == "Subject")
                {
                    SUBJECT = prt.Value;
                }

            }
            TITLE = FROM + " :: " + SUBJECT;





            String getNestedParts(IList<MessagePart> part, string curr)
            {
                string str = curr;
                if (part == null)
                {
                    return str;
                }
                else
                {
                    foreach (var parts in part)
                    {
                        if (parts.Parts == null)
                        {
                            if (parts.Body != null && parts.Body.Data != null)
                            {
                                String codedBody = parts.Body.Data.Replace("-", "+");
                                codedBody = codedBody.Replace("_", "/");
                                byte[] data = Convert.FromBase64String(codedBody);
                                codedBody = Encoding.UTF8.GetString(data);
                                str += codedBody;
                            }
                        }
                        else
                        {
                            return getNestedParts(parts.Parts, str);
                        }
                    }

                    return str;
                }



            }
        }




        



    }

}
