using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EDP_GUI
{
    // 1. EPPlus License Initializer
    public static class AppInitializer
    {
        public static void Init()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Student Project");
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AppInitializer.Init(); // Start Excel License
            Application.Run(new frmLogin());
        }
    }

    // 2. LOGIN FORM 
    public class frmLogin : Form
    {
        public frmLogin()
        {
            this.Text = "System Login";
            this.Size = new Size(350, 250);
            this.StartPosition = FormStartPosition.CenterScreen;

            Label lblTitle = new Label { Text = "Mecha Hobby Shop UI", Font = new Font("Segoe UI", 14, FontStyle.Bold), Location = new Point(60, 20), AutoSize = true };
            Label lblUser = new Label { Text = "Username:", Location = new Point(40, 70), AutoSize = true };
            TextBox txtUser = new TextBox { Location = new Point(120, 70), Width = 150 };
            Label lblPass = new Label { Text = "Password:", Location = new Point(40, 110), AutoSize = true };
            TextBox txtPass = new TextBox { Location = new Point(120, 110), Width = 150, PasswordChar = '*' };

            Button btnLogin = new Button { Text = "Login", Location = new Point(120, 150), Width = 70, BackColor = Color.LightBlue };
            btnLogin.Click += (s, e) => {
                try
                {
                    DatabaseConnection db = new DatabaseConnection();
                    using (MySqlConnection conn = db.GetConnection())
                    {
                        conn.Open();
                        string query = "SELECT * FROM users WHERE username = @user AND password = @pass AND account_status = 'Active'";
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@user", txtUser.Text);
                        cmd.Parameters.AddWithValue("@pass", txtPass.Text);
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                MessageBox.Show("Login Successful!");
                                new frmDashboard().Show();
                                this.Hide();
                            }
                            else
                            {
                                MessageBox.Show("Invalid Credentials or Account is Inactive.");
                            }
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Database Error. Is MySQL running?\n" + ex.Message); }
            };

            Button btnExit = new Button { Text = "Exit", Location = new Point(200, 150), Width = 70 };
            btnExit.Click += (s, e) => Application.Exit();

            LinkLabel lnkForgot = new LinkLabel { Text = "Forgot Password?", Location = new Point(120, 185), AutoSize = true };
            lnkForgot.LinkClicked += (s, e) => new frmRecovery().ShowDialog();

            this.Controls.AddRange(new Control[] { lblTitle, lblUser, txtUser, lblPass, txtPass, btnLogin, btnExit, lnkForgot });
        }
    }

    // 3. PASSWORD RECOVERY
    public class frmRecovery : Form
    {
        public frmRecovery()
        {
            this.Text = "Password Recovery";
            this.Size = new Size(400, 180);
            this.StartPosition = FormStartPosition.CenterParent;

            Label lblInfo = new Label { Text = "Enter your email to recover your password:", Location = new Point(20, 20), AutoSize = true };
            TextBox txtEmail = new TextBox { Location = new Point(20, 50), Width = 340 };

            Button btnSend = new Button { Text = "Recover", Location = new Point(20, 90), Width = 100, BackColor = Color.LightGreen };
            btnSend.Click += (s, e) => {
                try
                {
                    DatabaseConnection db = new DatabaseConnection();
                    using (MySqlConnection conn = db.GetConnection())
                    {
                        conn.Open();
                        string query = "SELECT password FROM users WHERE email = @email";
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                        object result = cmd.ExecuteScalar();

                        if (result != null) MessageBox.Show($"Your recovered password is: {result.ToString()}");
                        else MessageBox.Show("Email not found in the database.");
                    }
                }
                catch (Exception ex) { MessageBox.Show("Database Error: " + ex.Message); }
            };

            Button btnClose = new Button { Text = "Close", Location = new Point(130, 90), Width = 80 };
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { lblInfo, txtEmail, btnSend, btnClose });
        }
    }

    // 4. THE UPDATED DASHBOARD
    public class frmDashboard : Form
    {
        public frmDashboard()
        {
            this.Text = "Hobby Shop Dashboard - updated by Mizell";
            this.Size = new Size(600, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormClosed += (s, e) => Application.Exit();

            Label lblWelcome = new Label { Text = "Main Control Panel", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            // Links to Activity 5
            Button btnUsers = new Button { Text = "Manage Users", Location = new Point(20, 80), Size = new Size(160, 60), BackColor = Color.LightSkyBlue };
            btnUsers.Click += (s, e) => new frmUserManagement().ShowDialog();

            // Links to Activity 6 Transactions
            Button btnTransactions = new Button { Text = "Primary Transactions", Location = new Point(200, 80), Size = new Size(160, 60), BackColor = Color.LightGreen };
            btnTransactions.Click += (s, e) => new frmTransactions().ShowDialog();

            // Links to Activity 6 Reports
            Button btnReports = new Button { Text = "Generate Excel Reports", Location = new Point(380, 80), Size = new Size(160, 60), BackColor = Color.Orange };
            btnReports.Click += (s, e) => new frmReports().ShowDialog();

            Button btnLogout = new Button { Text = "Logout", Location = new Point(20, 280), Size = new Size(520, 40), BackColor = Color.LightCoral };
            btnLogout.Click += (s, e) => {
                new frmLogin().Show();
                this.Hide();
            };

            this.Controls.AddRange(new Control[] { lblWelcome, btnUsers, btnTransactions, btnReports, btnLogout });
        }
    }

    // 5. USER MANAGEMENT (From Activity 5)
    public class frmUserManagement : Form
    {
        public frmUserManagement()
        {
            this.Text = "User Management";
            this.Size = new Size(700, 450);
            this.StartPosition = FormStartPosition.CenterParent;

            Label lblUser = new Label { Text = "Username:", Location = new Point(20, 20), AutoSize = true };
            TextBox txtUser = new TextBox { Location = new Point(90, 17), Width = 120 };

            Label lblPass = new Label { Text = "Password:", Location = new Point(20, 50), AutoSize = true };
            TextBox txtPass = new TextBox { Location = new Point(90, 47), Width = 120 };

            Label lblEmail = new Label { Text = "Email:", Location = new Point(20, 80), AutoSize = true };
            TextBox txtEmail = new TextBox { Location = new Point(90, 77), Width = 120 };

            Label lblStatus = new Label { Text = "Status:", Location = new Point(20, 110), AutoSize = true };
            ComboBox cmbStatus = new ComboBox { Location = new Point(90, 107), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbStatus.Items.AddRange(new object[] { "Active", "Inactive" });
            cmbStatus.SelectedIndex = 0;

            DataGridView gridUsers = new DataGridView { Location = new Point(240, 50), Size = new Size(420, 330), ReadOnly = true };
            TextBox txtSearch = new TextBox { Location = new Point(240, 17), Width = 200 };
            Button btnSearch = new Button { Text = "Search User", Location = new Point(450, 15), Width = 100 };

            Button btnAdd = new Button { Text = "Add User", Location = new Point(20, 150), Width = 190, BackColor = Color.LightGreen };
            btnAdd.Click += (s, e) => {
                using (MySqlConnection conn = new DatabaseConnection().GetConnection())
                {
                    conn.Open();
                    string query = "INSERT INTO users (username, password, email, account_status) VALUES (@u, @p, @e, @st)";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@u", txtUser.Text);
                    cmd.Parameters.AddWithValue("@p", txtPass.Text);
                    cmd.Parameters.AddWithValue("@e", txtEmail.Text);
                    cmd.Parameters.AddWithValue("@st", cmbStatus.Text);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("User Added!");
                }
            };

            Button btnUpdate = new Button { Text = "Update Profile/Status", Location = new Point(20, 190), Width = 190, BackColor = Color.Orange };
            btnUpdate.Click += (s, e) => {
                using (MySqlConnection conn = new DatabaseConnection().GetConnection())
                {
                    conn.Open();
                    string query = "UPDATE users SET email=@e, account_status=@st WHERE username=@u";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@u", txtUser.Text);
                    cmd.Parameters.AddWithValue("@e", txtEmail.Text);
                    cmd.Parameters.AddWithValue("@st", cmbStatus.Text);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("User Updated!");
                }
            };

            btnSearch.Click += (s, e) => {
                using (MySqlConnection conn = new DatabaseConnection().GetConnection())
                {
                    conn.Open();
                    string query = "SELECT id, username, email, account_status FROM users WHERE username LIKE @search";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@search", "%" + txtSearch.Text + "%");
                    MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    gridUsers.DataSource = dt;
                }
            };

            this.Controls.AddRange(new Control[] { lblUser, txtUser, lblPass, txtPass, lblEmail, txtEmail, lblStatus, cmbStatus, btnAdd, btnUpdate, txtSearch, btnSearch, gridUsers });
        }
    }

    // 6. THE 3 PRIMARY TRANSACTIONS (New for Activity 6)
    public class frmTransactions : Form
    {
        public frmTransactions()
        {
            this.Text = "System Transactions";
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterParent;

            Label lbl = new Label { Text = "Select Transaction:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Arial", 12, FontStyle.Bold) };

            Button btnSale = new Button { Text = "1. Process Sale (HG-01 Gundam)", Location = new Point(20, 60), Size = new Size(300, 50) };
            btnSale.Click += (s, e) => ExecuteTransaction("INSERT INTO sales_transactions (item_code, qty_sold, total_amount) VALUES ('HG-01', 1, 1500.00)");

            Button btnRestock = new Button { Text = "2. Receive Inventory (Paints)", Location = new Point(20, 130), Size = new Size(300, 50) };
            btnRestock.Click += (s, e) => ExecuteTransaction("INSERT INTO inventory_restocks (item_code, qty_added, supplier) VALUES ('PT-01', 20, 'Tamiya Corp')");

            Button btnPreorder = new Button { Text = "3. Customer Pre-order", Location = new Point(20, 200), Size = new Size(300, 50) };
            btnPreorder.Click += (s, e) => ExecuteTransaction("INSERT INTO preorders (customer_name, item_name, deposit_amount) VALUES ('John Doe', 'PG Unleashed', 5000.00)");

            this.Controls.AddRange(new Control[] { lbl, btnSale, btnRestock, btnPreorder });
        }

        private void ExecuteTransaction(string query)
        {
            try
            {
                using (MySqlConnection conn = new DatabaseConnection().GetConnection())
                {
                    conn.Open();
                    new MySqlCommand(query, conn).ExecuteNonQuery();
                    MessageBox.Show("Transaction Logged Successfully in Database!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }
    }

    // 7. REPORT GENERATOR & EXCEL EXPORT (New for Activity 6)
    public class frmReports : Form
    {
        DataGridView gridReport;

        public frmReports()
        {
            this.Text = "Report Generator";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;

            Button btnLoad = new Button { Text = "Load Sales Data", Location = new Point(20, 20), Size = new Size(150, 40) };
            btnLoad.Click += BtnLoad_Click;

            Button btnExport = new Button { Text = "Export to Excel Template", Location = new Point(190, 20), Size = new Size(200, 40), BackColor = Color.Orange };
            btnExport.Click += BtnExport_Click;

            gridReport = new DataGridView { Location = new Point(20, 80), Size = new Size(640, 360), ReadOnly = true };

            this.Controls.AddRange(new Control[] { btnLoad, btnExport, gridReport });
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            try
            {
                using (MySqlConnection conn = new DatabaseConnection().GetConnection())
                {
                    conn.Open();
                    string query = "SELECT item_code, SUM(qty_sold) as Total_Qty, SUM(total_amount) as Revenue FROM sales_transactions GROUP BY item_code";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(new MySqlCommand(query, conn));
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    gridReport.DataSource = dt;
                }
            }
            catch (Exception ex) { MessageBox.Show("Database Error: " + ex.Message); }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (gridReport.Rows.Count == 0 || gridReport.DataSource == null)
            {
                MessageBox.Show("Please click 'Load Sales Data' first!"); return;
            }

            SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = "Sales_Report.xlsx" };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    // --- SHEET 1: DATA, LOGO & SIGNATURE ---
                    ExcelWorksheet ws1 = package.Workbook.Worksheets.Add("Report Data");

                    ws1.Cells["A1"].Value = "MECHA HOBBY SHOP INC.";
                    ws1.Cells["A1"].Style.Font.Size = 16;
                    ws1.Cells["A1"].Style.Font.Bold = true;
                    ws1.Cells["A2"].Value = "Official Sales Summary Report";
                    ws1.Cells["A3"].Value = "Generated on: " + DateTime.Now.ToString("yyyy-MM-dd");

                    DataTable dt = (DataTable)gridReport.DataSource;
                    ws1.Cells["A5"].LoadFromDataTable(dt, true);
                    ws1.Cells["A5:C5"].Style.Font.Bold = true;

                    int lastRow = dt.Rows.Count + 7;
                    ws1.Cells["A" + lastRow].Value = "Prepared and Verified By:";
                    ws1.Cells["A" + (lastRow + 2)].Value = "_______________________________";
                    ws1.Cells["A" + (lastRow + 3)].Value = "Authorized User Signature";

                    // --- SHEET 2: THE GRAPH ---
                    ExcelWorksheet ws2 = package.Workbook.Worksheets.Add("Data Visualization");
                    var chart = ws2.Drawings.AddChart("SalesChart", eChartType.ColumnClustered);
                    chart.Title.Text = "Revenue by Item Code";
                    chart.SetPosition(2, 0, 2, 0);
                    chart.SetSize(600, 400);

                    var yData = ws1.Cells[6, 3, dt.Rows.Count + 5, 3];
                    var xLabels = ws1.Cells[6, 1, dt.Rows.Count + 5, 1];
                    chart.Series.Add(yData, xLabels);

                    FileInfo fi = new FileInfo(sfd.FileName);
                    package.SaveAs(fi);
                }
                MessageBox.Show("Excel Report Generated Successfully! Check your saved file.");
            }
        }
    }
}
