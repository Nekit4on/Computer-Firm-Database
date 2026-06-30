using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class SuppliersForm : Form
    {
        private DataGridView dgvSuppliers;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnRefresh;
        private DataTable suppliersTable;

        public SuppliersForm()
        {
            InitializeComponents();
            LoadData();
        }

        private void InitializeComponents()
        {
            this.Text = "Управление поставщиками";
            this.Size = new Size(800, 450);
            this.StartPosition = FormStartPosition.CenterScreen;

            dgvSuppliers = new DataGridView()
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };

            Panel buttonPanel = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            btnAdd = new Button() { Text = "Добавить", Location = new Point(10, 10), Size = new Size(100, 30) };
            btnEdit = new Button() { Text = "Изменить", Location = new Point(120, 10), Size = new Size(100, 30) };
            btnDelete = new Button() { Text = "Удалить", Location = new Point(230, 10), Size = new Size(100, 30) };
            btnRefresh = new Button() { Text = "Обновить", Location = new Point(340, 10), Size = new Size(100, 30) };

            btnAdd.Click += BtnAdd_Click;
            btnEdit.Click += BtnEdit_Click;
            btnDelete.Click += BtnDelete_Click;
            btnRefresh.Click += (s, e) => LoadData();

            buttonPanel.Controls.Add(btnAdd);
            buttonPanel.Controls.Add(btnEdit);
            buttonPanel.Controls.Add(btnDelete);
            buttonPanel.Controls.Add(btnRefresh);

            this.Controls.Add(dgvSuppliers);
            this.Controls.Add(buttonPanel);
        }

        private void LoadData()
        {
            string query = "SELECT supplier_id, name, phone, address FROM suppliers ORDER BY supplier_id";
            suppliersTable = DbConnection.ExecuteQuery(query);
            dgvSuppliers.DataSource = suppliersTable;
            if (dgvSuppliers.Columns.Contains("supplier_id"))
                dgvSuppliers.Columns["supplier_id"].Visible = false;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string name = Microsoft.VisualBasic.Interaction.InputBox("Название поставщика:", "Добавление");
            if (string.IsNullOrEmpty(name)) return;
            string phone = Microsoft.VisualBasic.Interaction.InputBox("Телефон:", "Добавление");
            string address = Microsoft.VisualBasic.Interaction.InputBox("Адрес:", "Добавление");

            string query = "INSERT INTO suppliers (name, phone, address) VALUES (@name, @phone, @address)";
            NpgsqlParameter[] pars = {
                new NpgsqlParameter("@name", name),
                new NpgsqlParameter("@phone", phone),
                new NpgsqlParameter("@address", address)
            };
            DbConnection.ExecuteNonQuery(query, pars);
            LoadData();
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgvSuppliers.CurrentRow == null)
                return;

            int supplierId = Convert.ToInt32(
                dgvSuppliers.CurrentRow.Cells["supplier_id"].Value);

            string name =
                dgvSuppliers.CurrentRow.Cells["name"].Value?.ToString();

            string phone =
                dgvSuppliers.CurrentRow.Cells["phone"].Value?.ToString();

            string address =
                dgvSuppliers.CurrentRow.Cells["address"].Value?.ToString();

            Form editForm = new Form();

            editForm.Text = "Редактирование поставщика";
            editForm.Size = new Size(420, 260);
            editForm.StartPosition = FormStartPosition.CenterParent;
            editForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            editForm.MaximizeBox = false;

            Label lblName = new Label();
            lblName.Text = "Название:";
            lblName.Location = new Point(20, 30);
            lblName.AutoSize = true;

            TextBox txtName = new TextBox();
            txtName.Text = name;
            txtName.Location = new Point(140, 25);
            txtName.Size = new Size(220, 25);

            Label lblPhone = new Label();
            lblPhone.Text = "Телефон:";
            lblPhone.Location = new Point(20, 80);
            lblPhone.AutoSize = true;

            TextBox txtPhone = new TextBox();
            txtPhone.Text = phone;
            txtPhone.Location = new Point(140, 75);
            txtPhone.Size = new Size(220, 25);

            Label lblAddress = new Label();
            lblAddress.Text = "Адрес:";
            lblAddress.Location = new Point(20, 130);
            lblAddress.AutoSize = true;

            TextBox txtAddress = new TextBox();
            txtAddress.Text = address;
            txtAddress.Location = new Point(140, 125);
            txtAddress.Size = new Size(220, 25);

            Button btnSave = new Button();
            btnSave.Text = "Сохранить";
            btnSave.Size = new Size(120, 35);
            btnSave.Location = new Point(130, 175);

            btnSave.Click += (s, ev) =>
            {
                if (txtName.Text.Trim() == "")
                {
                    MessageBox.Show("Введите название поставщика");
                    return;
                }

                string query =
                    @"UPDATE suppliers
              SET name = @name,
                  phone = @phone,
                  address = @address
              WHERE supplier_id = @id";

                NpgsqlParameter[] pars =
                {
            new NpgsqlParameter("@name", txtName.Text),
            new NpgsqlParameter("@phone", txtPhone.Text),
            new NpgsqlParameter("@address", txtAddress.Text),
            new NpgsqlParameter("@id", supplierId)
        };

                DbConnection.ExecuteNonQuery(query, pars);

                LoadData();

                editForm.Close();
            };

            editForm.Controls.Add(lblName);
            editForm.Controls.Add(txtName);

            editForm.Controls.Add(lblPhone);
            editForm.Controls.Add(txtPhone);

            editForm.Controls.Add(lblAddress);
            editForm.Controls.Add(txtAddress);

            editForm.Controls.Add(btnSave);

            editForm.ShowDialog();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvSuppliers.CurrentRow == null) return;
            int id = Convert.ToInt32(dgvSuppliers.CurrentRow.Cells["supplier_id"].Value);
            if (MessageBox.Show("Удалить поставщика?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    DbConnection.ExecuteNonQuery("DELETE FROM suppliers WHERE supplier_id = @id", new[] { new NpgsqlParameter("@id", id) });
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Нельзя удалить поставщика, есть связанные комплектующие.");
                }
            }
        }
    }
}