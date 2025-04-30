using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Azure.Core;

namespace RIM
{
    public partial class Directory : Form
    {
        private string _accessToken;
        private DataTable _userTable;
        private DataTable _validationTable;
        private DataGridView dataGridViewUsers;

        public Directory()
        {
            InitializeComponent();
            _accessToken = accessToken;
            this.Load += Directory_Load;
        }


        private async void Directory_Load(object sender, EventArgs e)
        {
            try
            {
                _userTable = await LoadDirectoryUsers();
                AddMailboxStorageColumns(_userTable);
                await AddMailboxUsageData(_userTable);
                PrepareValidationData(_userTable);

                dataGridView.DataSource = _userTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Load error: " + ex.Message);
            }
        }
        private async void ActiveDirectory_Load(object sender, EventArgs e)
        {
            try
            {
                _userTable = await LoadDirectoryUsers();
                AddMailboxStorageColumns(_userTable);
                await AddMailboxUsageData(_userTable);
                PrepareValidationData(_userTable);

                dataGridView.DataSource = _userTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Load error: " + ex.Message);
            }
        }

        private async Task<DataTable> LoadDirectoryUsers()
        {
            var table = new DataTable();
            table.Columns.Add("Display Name");
            table.Columns.Add("Email");
            table.Columns.Add("Job Title");
            table.Columns.Add("Department");
            table.Columns.Add("Enabled");

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
                string url = "https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,jobTitle,department,accountEnabled";

                HttpResponseMessage response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject parsed = JObject.Parse(json);
                    JArray users = (JArray)parsed["value"];

                    foreach (var user in users)
                    {
                        table.Rows.Add(
                            (string)user["displayName"],
                            (string)user["userPrincipalName"],
                            (string)user["jobTitle"],
                            (string)user["department"],
                            user["accountEnabled"]?.ToString()
                        );
                    }
                }
                else
                {
                    MessageBox.Show("Graph error: " + response.StatusCode);
                }
            }
            return table;
        }
        private async Task AddMailboxUsageData(DataTable table)
        {
            await Task.Run(() =>
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddCommand("Connect-ExchangeOnline")
                      .AddParameter("ShowBanner", false);

                    try
                    {
                        ps.Invoke();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Exchange Online connection failed: " + ex.Message);
                        return;
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        string userEmail = row["Email"].ToString();

                        ps.Commands.Clear();
                        ps.AddCommand("Get-MailboxStatistics")
                          .AddParameter("Identity", userEmail);

                        try
                        {
                            var result = ps.Invoke();
                            if (ps.HadErrors || result.Count == 0)
                            {
                                row["Mailbox Total Size (MB)"] = "N/A";
                                row["Mailbox Quota (MB)"] = "N/A";
                                continue;
                            }

                            var stats = result[0];
                            var totalSize = stats.Members["TotalItemSize"].Value.ToString();
                            var quota = stats.Members["ProhibitSendQuota"].Value.ToString();

                            row["Mailbox Total Size (MB)"] = totalSize;
                            row["Mailbox Quota (MB)"] = quota;
                        }
                        catch
                        {
                            row["Mailbox Total Size (MB)"] = "Error";
                            row["Mailbox Quota (MB)"] = "Error";
                        }
                    }

                    ps.Commands.Clear();
                    ps.AddCommand("Disconnect-ExchangeOnline")
                      .AddParameter("Confirm", false);
                    ps.Invoke();
                }
            });
        private void PrepareValidationData(DataTable userTable)
        {
            _validationTable = userTable.Clone();
            foreach (DataRow row in userTable.Rows)
            {
                if (string.IsNullOrEmpty(row["Email"].ToString()) || row["Enabled"].ToString().ToLower() != "true")
                {
                    _validationTable.ImportRow(row);
                }
            }
        }


    }
}
