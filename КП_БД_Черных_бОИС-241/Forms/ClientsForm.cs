using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class ClientsForm : Form
    {
        private DataGridView dgvClients;
        private DataTable clientsTable;
        private Button btnAdd, btnEdit, btnDelete, btnRefresh, btnExcel, btnWord;

        public ClientsForm()
        {
            InitializeComponents();
            LoadData();
        }

        private void InitializeComponents()
        {
            this.Text = "Клиенты";
            this.Size = new Size(1100, 500);
            this.StartPosition = FormStartPosition.CenterScreen;

            dgvClients = new DataGridView()
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };
            this.Controls.Add(dgvClients);

            Panel btnPanel = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            btnAdd = new Button() { Text = "Добавить", Location = new Point(10, 10), Size = new Size(100, 30) };
            btnEdit = new Button() { Text = "Изменить", Location = new Point(120, 10), Size = new Size(100, 30) };
            btnDelete = new Button() { Text = "Удалить", Location = new Point(230, 10), Size = new Size(100, 30) };
            btnRefresh = new Button() { Text = "Обновить", Location = new Point(340, 10), Size = new Size(100, 30) };
            btnExcel = new Button() { Text = "Excel", Location = new Point(450, 10), Size = new Size(80, 30) };
            btnWord = new Button() { Text = "Word", Location = new Point(540, 10), Size = new Size(80, 30) };

            btnAdd.Click += BtnAdd_Click;
            btnEdit.Click += BtnEdit_Click;
            btnDelete.Click += BtnDelete_Click;
            btnRefresh.Click += (s, e) => LoadData();
            btnExcel.Click += (s, e) => ExportToExcel();
            btnWord.Click += (s, e) => ExportToWord();

            btnPanel.Controls.Add(btnAdd);
            btnPanel.Controls.Add(btnEdit);
            btnPanel.Controls.Add(btnDelete);
            btnPanel.Controls.Add(btnRefresh);
            btnPanel.Controls.Add(btnExcel);
            btnPanel.Controls.Add(btnWord);
            this.Controls.Add(btnPanel);
        }

        private void LoadData()
        {
            string query = "SELECT client_id, surname, name, patronymic, phone, email, address FROM clients ORDER BY client_id";
            clientsTable = DbConnection.ExecuteQuery(query);
            dgvClients.DataSource = clientsTable;
            if (dgvClients.Columns.Contains("client_id"))
                dgvClients.Columns["client_id"].Visible = false;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string surname = Microsoft.VisualBasic.Interaction.InputBox("Фамилия:");
            if (string.IsNullOrEmpty(surname)) return;
            string name = Microsoft.VisualBasic.Interaction.InputBox("Имя:");
            string patronymic = Microsoft.VisualBasic.Interaction.InputBox("Отчество:");
            string phone = Microsoft.VisualBasic.Interaction.InputBox("Телефон:");
            string email = Microsoft.VisualBasic.Interaction.InputBox("Email:");
            string address = Microsoft.VisualBasic.Interaction.InputBox("Адрес:");

            string query = @"INSERT INTO clients (surname, name, patronymic, phone, email, address)
                             VALUES (@surname, @name, @patronymic, @phone, @email, @address)";
            NpgsqlParameter[] pars = {
                new NpgsqlParameter("@surname", surname),
                new NpgsqlParameter("@name", name),
                new NpgsqlParameter("@patronymic", string.IsNullOrEmpty(patronymic) ? (object)DBNull.Value : patronymic),
                new NpgsqlParameter("@phone", phone),
                new NpgsqlParameter("@email", email),
                new NpgsqlParameter("@address", address)
            };
            DbConnection.ExecuteNonQuery(query, pars);
            LoadData();
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgvClients.CurrentRow == null) return;
            int id = Convert.ToInt32(dgvClients.CurrentRow.Cells["client_id"].Value);
            string oldSurname = dgvClients.CurrentRow.Cells["surname"].Value.ToString();
            string oldName = dgvClients.CurrentRow.Cells["name"].Value.ToString();
            string oldPatronymic = dgvClients.CurrentRow.Cells["patronymic"].Value?.ToString();
            string oldPhone = dgvClients.CurrentRow.Cells["phone"].Value?.ToString();
            string oldEmail = dgvClients.CurrentRow.Cells["email"].Value?.ToString();
            string oldAddress = dgvClients.CurrentRow.Cells["address"].Value?.ToString();

            Form dialog = new Form() { Text = "Редактирование клиента", Size = new Size(400, 350), StartPosition = FormStartPosition.CenterParent };
            TextBox txtSurname = new TextBox() { Text = oldSurname, Location = new Point(120, 20), Width = 200 };
            TextBox txtName = new TextBox() { Text = oldName, Location = new Point(120, 60), Width = 200 };
            TextBox txtPatronymic = new TextBox() { Text = oldPatronymic, Location = new Point(120, 100), Width = 200 };
            TextBox txtPhone = new TextBox() { Text = oldPhone, Location = new Point(120, 140), Width = 200 };
            TextBox txtEmail = new TextBox() { Text = oldEmail, Location = new Point(120, 180), Width = 200 };
            TextBox txtAddress = new TextBox() { Text = oldAddress, Location = new Point(120, 220), Width = 200 };

            dialog.Controls.Add(new Label() { Text = "Фамилия:", Location = new Point(20, 22) }); dialog.Controls.Add(txtSurname);
            dialog.Controls.Add(new Label() { Text = "Имя:", Location = new Point(20, 62) }); dialog.Controls.Add(txtName);
            dialog.Controls.Add(new Label() { Text = "Отчество:", Location = new Point(20, 102) }); dialog.Controls.Add(txtPatronymic);
            dialog.Controls.Add(new Label() { Text = "Телефон:", Location = new Point(20, 142) }); dialog.Controls.Add(txtPhone);
            dialog.Controls.Add(new Label() { Text = "Email:", Location = new Point(20, 182) }); dialog.Controls.Add(txtEmail);
            dialog.Controls.Add(new Label() { Text = "Адрес:", Location = new Point(20, 222) }); dialog.Controls.Add(txtAddress);

            Button btnSave = new Button() { Text = "Сохранить", Location = new Point(100, 270), Size = new Size(100, 30) };
            btnSave.Click += (s, ev) =>
            {
                string query = @"UPDATE clients SET surname=@s, name=@n, patronymic=@p, phone=@ph, email=@e, address=@a WHERE client_id=@id";
                NpgsqlParameter[] pars = {
                    new NpgsqlParameter("@s", txtSurname.Text),
                    new NpgsqlParameter("@n", txtName.Text),
                    new NpgsqlParameter("@p", string.IsNullOrEmpty(txtPatronymic.Text) ? (object)DBNull.Value : txtPatronymic.Text),
                    new NpgsqlParameter("@ph", txtPhone.Text),
                    new NpgsqlParameter("@e", txtEmail.Text),
                    new NpgsqlParameter("@a", txtAddress.Text),
                    new NpgsqlParameter("@id", id)
                };
                DbConnection.ExecuteNonQuery(query, pars);
                LoadData();
                dialog.Close();
            };
            dialog.Controls.Add(btnSave);
            dialog.ShowDialog();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvClients.CurrentRow == null) return;
            int id = Convert.ToInt32(dgvClients.CurrentRow.Cells["client_id"].Value);
            if (MessageBox.Show("Удалить клиента?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DbConnection.ExecuteNonQuery("DELETE FROM clients WHERE client_id = @id", new[] { new NpgsqlParameter("@id", id) });
                LoadData();
            }
        }

        private void ExportToExcel()
        {
            if (dgvClients.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excel = Activator.CreateInstance(excelType);
                excel.Visible = true;
                var workbook = excel.Workbooks.Add();
                var sheet = workbook.ActiveSheet;
                for (int i = 0; i < dgvClients.Columns.Count; i++)
                    sheet.Cells[1, i + 1] = dgvClients.Columns[i].HeaderText;
                for (int i = 0; i < dgvClients.Rows.Count; i++)
                    for (int j = 0; j < dgvClients.Columns.Count; j++)
                        sheet.Cells[i + 2, j + 1] = dgvClients.Rows[i].Cells[j].Value?.ToString();
                MessageBox.Show("Экспорт в Excel выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Excel: " + ex.Message); }
        }

        private void ExportToWord()
        {
            if (dgvClients.Rows.Count == 0) return;
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                dynamic word = Activator.CreateInstance(wordType);
                word.Visible = true;
                var doc = word.Documents.Add();
                int rows = dgvClients.Rows.Count;
                int cols = dgvClients.Columns.Count;
                var table = doc.Tables.Add(doc.Range(), rows, cols);
                for (int i = 0; i < cols; i++)
                    table.Cell(1, i + 1).Range.Text = dgvClients.Columns[i].HeaderText;
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                        table.Cell(i + 2, j + 1).Range.Text = dgvClients.Rows[i].Cells[j].Value?.ToString() ?? "";
                MessageBox.Show("Экспорт в Word выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Word: " + ex.Message); }
        }
    }
}