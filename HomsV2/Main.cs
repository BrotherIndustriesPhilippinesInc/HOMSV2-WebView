using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HomsV2.Forms;
using HomsV2.Functions;
using Newtonsoft.Json.Linq;
using System.IO;
using WMPLib;
using Newtonsoft.Json;
using System.Diagnostics;

namespace HomsV2
{
    public partial class Main: Form
    {
        // Constants
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // Import user32.dll for sending messages
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        public static string localIP;
        public static string UserIdNumber = "";

        private string centralizedLoginConnString = "Data Source=APBIPHBPSDB02;Initial Catalog=Centralized_LOGIN_DB;Persist Security Info=True;User ID=CAS_access;Password=@BIPH2024"; 
        public static string GetLocalIPAddress()
        {
            foreach (var ip in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }

            throw new Exception("No network adapters with an IPv4 address in the system!");
        }

        private void GetIPAddressFromCentralizedLogin()
        {
            SqlConnection CentralizedLogin = new SqlConnection(centralizedLoginConnString);

            CentralizedLogin.Open();
            SqlCommand SelectUserAccount = new SqlCommand("SP_SelectLoginRequestFromCentralizedLogin", CentralizedLogin);
            SelectUserAccount.CommandType = CommandType.StoredProcedure;
            SelectUserAccount.Parameters.AddWithValue("@IPAddress", localIP);
            SelectUserAccount.Parameters.AddWithValue("@SystemID", "64"); //palitan ng assigned system id
            SqlDataAdapter da = new SqlDataAdapter(SelectUserAccount);
            DataTable dt = new DataTable();
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                SqlDataReader reader = SelectUserAccount.ExecuteReader();

                if (reader.Read())
                {
                    UserIdNumber = reader["USERNAME"].ToString(); //ito ay depende kong anong gamit nyo na user name (ADID/ID number)
                }
            }
            else
            {
                
                LoginReminders loginReminders = new LoginReminders();
                loginReminders.ShowDialog();
                return;
            }
        }

        private APIHandler apiHandler;

        private WebViewFunctions webViewFunctions;

        private SoundManager soundManager;

        private System.Windows.Forms.Timer linestopTimer;

        public Main()
        {
            localIP = GetLocalIPAddress();
            GetIPAddressFromCentralizedLogin();

            InitializeComponent();

            apiHandler = new APIHandler();
            webViewFunctions = new WebViewFunctions(webView21, CustomEnvironment: true);

            soundManager = new SoundManager();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            exit.ForeColor = Color.Red;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            exit.ForeColor = Color.White;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (e.Clicks == 2) // Detect double-click
                {
                    this.WindowState = this.WindowState == FormWindowState.Normal
                        ? FormWindowState.Maximized
                        : FormWindowState.Normal;
                }
                else
                {
                    // Handle dragging the form
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
                }
            }
        }

        private async void webView21_CoreWebView2InitializationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            await webViewFunctions.LoadPageAsync("http://apbiphbpswb02/homs/");

            if (e.IsSuccess)
            {
                // Allow the popup to happen naturally. 
                // Because they share the same 'CoreWebView2Environment', 
                // the cookies from the popup WILL be available to the main viz.
                webView21.CoreWebView2.NewWindowRequested += (s, args) =>
                {
                    // If you want it to open in a new popup window (cleanest for login):
                    // Just leave args.Handled = false (default) or don't handle this event at all.

                    // If the popup is definitely a login redirect:
                    Debug.WriteLine($"Popup requested for: {args.Uri}");
                };
            }

            Dictionary<string, string> post = new Dictionary<string, string> {
                { "id_number", UserIdNumber.ToString() }
            };
            JObject data = await apiHandler.APIPostCall("http://apbiphbpswb02/homs/api/user/getUser.php", post);

            await webViewFunctions.ExecuteJavascript($"localStorage.setItem(\"user\", JSON.stringify({data["data"]}));");

            await webViewFunctions.ExecuteJavascript($@"
                fetch('/homs/helpers/set_user.php', {{
                    method: 'POST',
                    headers: {{
                        'Content-Type': 'application/json'
                    }},
                    body: JSON.stringify({data["data"]})
                }});
                window.location.reload();
            ");


            string section = data["data"]["Section"]?.ToString();

            // Initialize and start the timer
            linestopTimer = new System.Windows.Forms.Timer();
            linestopTimer.Interval = 2000; // 1 second
            linestopTimer.Tick += async (s, args) =>
            {
                await CheckLinestops(section);
            };
            linestopTimer.Start();
        }

        private async void webView21_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            
        }

        private void maximize_Click(object sender, EventArgs e)
        {
            this.WindowState = this.WindowState == FormWindowState.Normal
               ? FormWindowState.Maximized
               : FormWindowState.Normal;
        }

        private void minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
             
        }

        private void webView21_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            JObject data = webViewFunctions.CapturedMessage(e);
            string jsonString = data.ToString(Newtonsoft.Json.Formatting.Indented);

            if (data["actionName"].ToString() == "generateDPR")
            {
                JArray prod = data["data"]["production_data"] as JArray;
                JObject wc = data["data"]["workcenter_data"] as JObject;
                JObject user = data["data"]["user_data"] as JObject;

                DPR dpr = new DPR(prod, wc, user);
                string generatedFilePath = dpr.generateDPRFile();
                JObject fileDetails = new JObject();
                fileDetails["generatedFilePath"] = generatedFilePath;

                webViewFunctions.SendDataToWeb(fileDetails, "generatedFilePath");
            }
            else if(data["actionName"].ToString() == "dismissNotification")
            {
                soundManager.StopSound();
            }
            else if (data["actionName"].ToString() == "checkDelayButtonClicked")
            {
                CheckDelay checkDelay = new CheckDelay();
                checkDelay.Show();
            }
        }

        private async Task CheckLinestops(string section)
        {
            // Call API for linestops of the section for the last 10 seconds
            APIHandler apiHandler = new APIHandler();
            JObject json = await apiHandler.APIGetCall($"http://apbiphbpswb02/homs/api/production/getRecentLinestops.php?section={section}");

            // Extract the array from the "data" property
            JArray result = json["data"] as JArray;

            if (result == null || !result.Any())
                return; // No linestops found, exit early

            foreach (var item in result)
            {
                int production_id;
                string timeCreated = item["time_updated"]?.ToString();
                if (!string.IsNullOrEmpty(timeCreated) && IsRecent(timeCreated))
                {
                    production_id = Convert.ToInt32(item["id"]);

                    SendLinestopDetailsToWeb(production_id);

                    string soundFile = "linestop.mp3";
                    string soundPath = Path.Combine(Application.StartupPath, "Resources", "Sounds", soundFile);
                    soundManager.PlaySound(soundPath, true, 20);
                }
            }
        }

        private void SendLinestopDetailsToWeb(int production_id)
        {
            JObject linestopDetails = new JObject(){{"production_id", production_id}};
            webViewFunctions.SendDataToWeb(linestopDetails, "linestop_details");
        }       

        private bool IsRecent(string timestampString, int thresholdInSeconds = 2)
        {
            // Parse timestamp (replace space with 'T' for ISO format)
            if (DateTime.TryParse(timestampString.Replace(" ", "T"), out DateTime entryTime))
            {
                DateTime now = DateTime.Now;
                double diffInSeconds = (now - entryTime).TotalSeconds;
                return diffInSeconds <= thresholdInSeconds;
            }

            // If parsing fails, return false
            return false;
        }
    }
}