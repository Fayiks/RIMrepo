using System.Windows.Forms;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace RIM
{
    public partial class Form1 : Form
    {
        private IPublicClientApplication _pca;
        private string[] _scopes = new[] { "User.Read", "User.Read.All", "Directory.Read.All" };

        public Form1()
        {
            InitializeComponent();

            _pca = PublicClientApplicationBuilder.Create("374b0ee8-30fb-4521-ba42-ea6279ddb271")
                .WithAuthority(AzureCloudInstance.AzurePublic, "f2b48808-df47-4e2f-b8fc-0a338e0dc3de") // or "common"
                .WithRedirectUri("http://localhost")
                .Build();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Signing in...");

            try
            {
                var result = await _pca.AcquireTokenInteractive(_scopes).ExecuteAsync();
                MessageBox.Show($"Welcome, {result.Account.Username}");



            }
            catch (MsalException ex)
            {
                MessageBox.Show($"Auth failed: {ex.Message}");
            }
        }
    }
}
