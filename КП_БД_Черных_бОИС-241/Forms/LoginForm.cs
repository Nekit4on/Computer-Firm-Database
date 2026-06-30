using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class LoginForm : Form
    {
        private TextBox txtLogin;
        private TextBox txtPassword;
        private Button btnLogin;

        public LoginForm()
        {
            InitializeComponents();
            this.AcceptButton = btnLogin;
        }

        private void InitializeComponents()
        {
            this.Text = "Вход в систему";
            this.Size = new Size(350, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            Label lblLogin = new Label()
            {
                Text = "Логин:",
                Location = new Point(30, 30),
                Size = new Size(60, 25)
            };

            txtLogin = new TextBox()
            {
                Name = "txtLogin",
                Location = new Point(100, 30),
                Size = new Size(180, 25)
            };

            Label lblPassword = new Label()
            {
                Text = "Пароль:",
                Location = new Point(30, 70),
                Size = new Size(60, 25)
            };

            txtPassword = new TextBox()
            {
                Name = "txtPassword",
                Location = new Point(100, 70),
                Size = new Size(180, 25),
                PasswordChar = '*'
            };

            btnLogin = new Button()
            {
                Name = "btnLogin",
                Text = "Войти",
                Location = new Point(130, 110),
                Size = new Size(75, 30)
            };
            btnLogin.Click += BtnLogin_Click;

            this.Controls.Add(lblLogin);
            this.Controls.Add(txtLogin);
            this.Controls.Add(lblPassword);
            this.Controls.Add(txtPassword);
            this.Controls.Add(btnLogin);
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text.Trim();

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Введите логин и пароль!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string query = @"
                SELECT u.user_id, u.employee_id, u.role, e.full_name 
                FROM users u
                JOIN employees e ON u.employee_id = e.employee_id
                WHERE u.login = @login AND u.password_hash = @pwd";

            NpgsqlParameter[] parameters = {
                new NpgsqlParameter("@login", login),
                new NpgsqlParameter("@pwd", password)
            };

            DataTable result = DbConnection.ExecuteQuery(query, parameters);

            if (result.Rows.Count == 1)
            {
                UserSession.UserId = Convert.ToInt32(result.Rows[0]["user_id"]);
                UserSession.EmployeeId = Convert.ToInt32(result.Rows[0]["employee_id"]);
                UserSession.Role = result.Rows[0]["role"].ToString();
                UserSession.FullName = result.Rows[0]["full_name"].ToString();

                this.Hide();
                MainForm main = new MainForm();
                main.FormClosed += (s, args) => this.Close();
                main.Show();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль.", "Ошибка авторизации", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Clear();
                txtPassword.Focus();
            }
        }
    }

    public static class UserSession
    {
        public static int UserId { get; set; }
        public static int EmployeeId { get; set; }
        public static string Role { get; set; }
        public static string FullName { get; set; }
    }
}