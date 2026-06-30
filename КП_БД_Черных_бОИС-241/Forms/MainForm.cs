using System;
using System.Drawing;
using System.Windows.Forms;

namespace КП_БД_Черных.Forms
{
    public partial class MainForm : Form
    {
        private Button btnClients;
        private Button btnComponents;
        private Button btnOrders;
        private Button btnSuppliers;
        private Button btnReports;
        private Button btnLogout;
        private Label lblUser;
        private Button currentActiveButton;

        public MainForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Компьютерная фирма - Главное меню";
            this.Size = new Size(800, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.IsMdiContainer = true;

            Panel menuPanel = new Panel()
            {
                Dock = DockStyle.Left,
                Width = 200,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            lblUser = new Label()
            {
                Text = $"Пользователь: {UserSession.FullName} ({UserSession.Role})",
                Dock = DockStyle.Top,
                Height = 40,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White
            };

            btnClients = CreateMenuButton("Клиенты", 1);
            btnComponents = CreateMenuButton("Комплектующие", 2);
            btnOrders = CreateMenuButton("Заказы", 3);
            btnSuppliers = CreateMenuButton("Поставщики", 4);
            btnReports = CreateMenuButton("Отчёты", 5);
            btnLogout = CreateMenuButton("Выход", 6);

            btnClients.Click += (s, e) => OpenForm(new ClientsForm(), btnClients);
            btnComponents.Click += (s, e) => OpenForm(new ComponentsForm(), btnComponents);
            btnOrders.Click += (s, e) => OpenForm(new OrdersForm(), btnOrders);
            btnSuppliers.Click += (s, e) => OpenForm(new SuppliersForm(), btnSuppliers);
            btnReports.Click += (s, e) => OpenForm(new ReportsForm(), btnReports);
            btnLogout.Click += (s, e) => Logout();

            // Разграничение прав
            switch (UserSession.Role)
            {
                case "Admin":
                    break;
                case "Manager":
                    btnComponents.Enabled = false;
                    btnSuppliers.Enabled = false;
                   
                    break;
                case "Stockman":
                    btnClients.Enabled = false;
                    btnOrders.Enabled = false;
                    btnReports.Enabled = false;
                    btnSuppliers.Enabled = false;
                    break;
            }

            menuPanel.Controls.Add(lblUser);
            menuPanel.Controls.Add(btnClients);
            menuPanel.Controls.Add(btnComponents);
            menuPanel.Controls.Add(btnOrders);
            menuPanel.Controls.Add(btnSuppliers);
            menuPanel.Controls.Add(btnReports);
            menuPanel.Controls.Add(btnLogout);

            this.Controls.Add(menuPanel);
        }

        private Button CreateMenuButton(string text, int order)
        {
            return new Button()
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 50,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.White,
                FlatAppearance = { BorderSize = 0 }
            };
        }

        private void OpenForm(Form form, Button activeButton)
        {
            form.MdiParent = this;
            form.WindowState = FormWindowState.Maximized;
            form.Show();
            if (currentActiveButton != null)
            {
                currentActiveButton.BackColor = Color.White;
                currentActiveButton.ForeColor = Color.Black;
            }
            currentActiveButton = activeButton;
            currentActiveButton.BackColor = Color.FromArgb(0, 120, 215);
            currentActiveButton.ForeColor = Color.White;
        }

        private void Logout()
        {
            DialogResult res = MessageBox.Show("Вы действительно хотите выйти?", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                Application.Restart();
            }
        }
    }
}