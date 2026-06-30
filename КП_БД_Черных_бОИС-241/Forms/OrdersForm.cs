using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class OrdersForm : Form
    {
        private DataGridView dgvOrders;
        private DataTable ordersTable;
        private Button btnNew, btnEdit, btnDelete, btnRefresh, btnExcel, btnWord, btnCsv;

        public OrdersForm()
        {
            InitializeComponents();
            LoadOrders();
        }

        private void InitializeComponents()
        {
            this.Text = "Заказы";
            this.Size = new Size(1100, 500);
            this.StartPosition = FormStartPosition.CenterScreen;

            dgvOrders = new DataGridView()
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };
            this.Controls.Add(dgvOrders);

            Panel btnPanel = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            btnNew = new Button() { Text = "Новый заказ", Location = new Point(10, 10), Size = new Size(110, 30) };
            btnEdit = new Button() { Text = "Редактировать", Location = new Point(130, 10), Size = new Size(110, 30) };
            btnDelete = new Button() { Text = "Удалить заказ", Location = new Point(250, 10), Size = new Size(110, 30) };
            btnRefresh = new Button() { Text = "Обновить", Location = new Point(370, 10), Size = new Size(100, 30) };
            btnExcel = new Button() { Text = "Excel", Location = new Point(480, 10), Size = new Size(80, 30) };
            btnWord = new Button() { Text = "Word", Location = new Point(570, 10), Size = new Size(80, 30) };
            btnCsv = new Button() { Text = "CSV", Location = new Point(660, 10), Size = new Size(80, 30) };

            btnNew.Click += BtnNew_Click;
            btnEdit.Click += BtnEdit_Click;
            btnDelete.Click += BtnDelete_Click;
            btnRefresh.Click += (s, e) => LoadOrders();
            btnExcel.Click += (s, e) => ExportToExcel();
            btnWord.Click += (s, e) => ExportToWord();
            btnCsv.Click += (s, e) => ExportToCsv();

            btnPanel.Controls.Add(btnNew);
            btnPanel.Controls.Add(btnEdit);
            btnPanel.Controls.Add(btnDelete);
            btnPanel.Controls.Add(btnRefresh);
            btnPanel.Controls.Add(btnExcel);
            btnPanel.Controls.Add(btnWord);
            btnPanel.Controls.Add(btnCsv);
            this.Controls.Add(btnPanel);
        }

        private void LoadOrders()
        {
            string query = @"
                SELECT o.order_id, o.order_date, 
                       c.surname || ' ' || c.name as client, 
                       e.full_name as manager, 
                       o.status, o.total_amount, o.payment,
                       o.delivery_date,
                       o.client_id, o.manager_id
                FROM orders o
                JOIN clients c ON o.client_id = c.client_id
                JOIN employees e ON o.manager_id = e.employee_id
                ORDER BY o.order_id DESC";
            ordersTable = DbConnection.ExecuteQuery(query);
            dgvOrders.DataSource = ordersTable;
            if (dgvOrders.Columns.Contains("order_id")) dgvOrders.Columns["order_id"].Visible = false;
            if (dgvOrders.Columns.Contains("client_id")) dgvOrders.Columns["client_id"].Visible = false;
            if (dgvOrders.Columns.Contains("manager_id")) dgvOrders.Columns["manager_id"].Visible = false;
        }

        private void BtnNew_Click(object sender, EventArgs e) { EditOrderDialog(0); }
        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgvOrders.CurrentRow == null) return;
            int orderId = Convert.ToInt32(dgvOrders.CurrentRow.Cells["order_id"].Value);
            EditOrderDialog(orderId);
        }
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvOrders.CurrentRow == null) return;
            int orderId = Convert.ToInt32(dgvOrders.CurrentRow.Cells["order_id"].Value);
            if (MessageBox.Show("Удалить заказ и все его позиции?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DbConnection.ExecuteNonQuery("DELETE FROM orders WHERE order_id = @oid", new[] { new NpgsqlParameter("@oid", orderId) });
                LoadOrders();
            }
        }

        private void EditOrderDialog(int orderId)
        {
            Form dialog = new Form();
            dialog.Text = orderId == 0 ? "Новый заказ" : "Редактирование заказа";
            dialog.Size = new Size(900, 650);
            dialog.StartPosition = FormStartPosition.CenterParent;

            DataTable clients = DbConnection.ExecuteQuery("SELECT client_id, surname || ' ' || name as fullname FROM clients");
            ComboBox cmbClient = new ComboBox() { DataSource = clients, DisplayMember = "fullname", ValueMember = "client_id", Location = new Point(120, 20), Width = 200 };
            Label lblClient = new Label() { Text = "Клиент:", Location = new Point(20, 22) };

            DataTable managers = DbConnection.ExecuteQuery("SELECT employee_id, full_name FROM employees");
            ComboBox cmbManager = new ComboBox() { DataSource = managers, DisplayMember = "full_name", ValueMember = "employee_id", Location = new Point(120, 60), Width = 200 };
            Label lblManager = new Label() { Text = "Менеджер:", Location = new Point(20, 62) };

            ComboBox cmbStatus = new ComboBox() { Items = { "New", "Processing", "Assembled", "Delivered", "Cancelled" }, Location = new Point(120, 100), Width = 150 };
            Label lblStatus = new Label() { Text = "Статус:", Location = new Point(20, 102) };

            TextBox txtPayment = new TextBox() { Location = new Point(120, 140), Width = 200 };
            Label lblPayment = new Label() { Text = "Способ оплаты:", Location = new Point(20, 142) };

            DateTimePicker dtpDelivery = new DateTimePicker() { Format = DateTimePickerFormat.Short, Location = new Point(120, 180), Width = 120 };
            Label lblDelivery = new Label() { Text = "Дата поставки:", Location = new Point(20, 182) };

            DataGridView dgvItems = new DataGridView()
            {
                Location = new Point(20, 220),
                Size = new Size(840, 250),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };

            Button btnAddItem = new Button() { Text = "Добавить позицию", Location = new Point(20, 480), Size = new Size(130, 30) };
            Button btnRemoveItem = new Button() { Text = "Удалить позицию", Location = new Point(160, 480), Size = new Size(130, 30) };
            Button btnSaveOrder = new Button() { Text = "Сохранить заказ", Location = new Point(320, 480), Size = new Size(130, 30) };
            Button btnCancel = new Button() { Text = "Отмена", Location = new Point(460, 480), Size = new Size(100, 30) };

            dialog.Controls.Add(lblClient); dialog.Controls.Add(cmbClient);
            dialog.Controls.Add(lblManager); dialog.Controls.Add(cmbManager);
            dialog.Controls.Add(lblStatus); dialog.Controls.Add(cmbStatus);
            dialog.Controls.Add(lblPayment); dialog.Controls.Add(txtPayment);
            dialog.Controls.Add(lblDelivery); dialog.Controls.Add(dtpDelivery);
            dialog.Controls.Add(dgvItems);
            dialog.Controls.Add(btnAddItem); dialog.Controls.Add(btnRemoveItem);
            dialog.Controls.Add(btnSaveOrder); dialog.Controls.Add(btnCancel);

            DataTable itemsTable = null;
            void LoadItems(int oid)
            {
                string query = @"
                    SELECT oi.configuration_id, c.name as config_name, oi.quantity, oi.unit_price, 
                           (oi.quantity * oi.unit_price) as total
                    FROM order_items oi
                    JOIN configurations c ON oi.configuration_id = c.configuration_id
                    WHERE oi.order_id = @oid";
                itemsTable = DbConnection.ExecuteQuery(query, new[] { new NpgsqlParameter("@oid", oid) });
                dgvItems.DataSource = itemsTable;
                if (dgvItems.Columns.Contains("configuration_id")) dgvItems.Columns["configuration_id"].Visible = false;
            }

            if (orderId != 0)
            {
                DataRow[] rows = ordersTable.Select($"order_id = {orderId}");
                if (rows.Length > 0)
                {
                    DataRow row = rows[0];
                    cmbClient.SelectedValue = row["client_id"];
                    cmbManager.SelectedValue = row["manager_id"];
                    cmbStatus.SelectedItem = row["status"].ToString();
                    txtPayment.Text = row["payment"].ToString();
                    dtpDelivery.Value = row["delivery_date"] == DBNull.Value ? DateTime.Today : Convert.ToDateTime(row["delivery_date"]);
                    LoadItems(orderId);
                }
            }

            btnAddItem.Click += (s, ev) =>
            {
                if (orderId == 0) { MessageBox.Show("Сначала сохраните заказ."); return; }
                DataTable configs = DbConnection.ExecuteQuery("SELECT configuration_id, name FROM configurations");
                Form configDialog = new Form() { Text = "Выбор конфигурации", Size = new Size(400, 150), StartPosition = FormStartPosition.CenterParent };
                ComboBox cmbConfig = new ComboBox() { DataSource = configs, DisplayMember = "name", ValueMember = "configuration_id", Location = new Point(20, 20), Width = 250 };
                NumericUpDown nudQty = new NumericUpDown() { Location = new Point(20, 60), Minimum = 1, Maximum = 100, Value = 1, Width = 80 };
                Button ok = new Button() { Text = "OK", Location = new Point(120, 55), Width = 80 };
                ok.Click += (ok_s, ok_e) =>
                {
                    int cfgId = (int)cmbConfig.SelectedValue;
                    int qty = (int)nudQty.Value;
                    string priceQuery = @"
                        SELECT SUM(cc.quantity * c.retail_price) as total_price
                        FROM configuration_components cc
                        JOIN components c ON cc.component_id = c.component_id
                        WHERE cc.configuration_id = @cfgId";
                    DataTable priceDt = DbConnection.ExecuteQuery(priceQuery, new[] { new NpgsqlParameter("@cfgId", cfgId) });
                    decimal unitPrice = priceDt.Rows[0]["total_price"] == DBNull.Value ? 0 : Convert.ToDecimal(priceDt.Rows[0]["total_price"]);
                    if (unitPrice == 0) { MessageBox.Show("Не удалось рассчитать цену"); return; }
                    string ins = "INSERT INTO order_items (order_id, configuration_id, quantity, unit_price) VALUES (@oid, @cfg, @qty, @price)";
                    DbConnection.ExecuteNonQuery(ins, new[] { new NpgsqlParameter("@oid", orderId), new NpgsqlParameter("@cfg", cfgId), new NpgsqlParameter("@qty", qty), new NpgsqlParameter("@price", unitPrice) });
                    LoadItems(orderId);
                    string sumQuery = "SELECT COALESCE(SUM(quantity * unit_price), 0) FROM order_items WHERE order_id = @oid";
                    DataTable sumDt = DbConnection.ExecuteQuery(sumQuery, new[] { new NpgsqlParameter("@oid", orderId) });
                    decimal newTotal = Convert.ToDecimal(sumDt.Rows[0][0]);
                    DbConnection.ExecuteNonQuery("UPDATE orders SET total_amount = @total WHERE order_id = @oid", new[] { new NpgsqlParameter("@total", newTotal), new NpgsqlParameter("@oid", orderId) });
                    configDialog.Close();
                    MessageBox.Show("Позиция добавлена.");
                };
                configDialog.Controls.Add(cmbConfig); configDialog.Controls.Add(nudQty); configDialog.Controls.Add(ok);
                configDialog.ShowDialog();
            };

            btnRemoveItem.Click += (s, ev) =>
            {
                if (dgvItems.CurrentRow == null) return;
                if (orderId == 0) { MessageBox.Show("Заказ не сохранён."); return; }
                int cfgId = Convert.ToInt32(dgvItems.CurrentRow.Cells["configuration_id"].Value);
                DbConnection.ExecuteNonQuery("DELETE FROM order_items WHERE order_id = @oid AND configuration_id = @cfg", new[] { new NpgsqlParameter("@oid", orderId), new NpgsqlParameter("@cfg", cfgId) });
                LoadItems(orderId);
                string sumQuery = "SELECT COALESCE(SUM(quantity * unit_price), 0) FROM order_items WHERE order_id = @oid";
                DataTable sumDt = DbConnection.ExecuteQuery(sumQuery, new[] { new NpgsqlParameter("@oid", orderId) });
                decimal newTotal = Convert.ToDecimal(sumDt.Rows[0][0]);
                DbConnection.ExecuteNonQuery("UPDATE orders SET total_amount = @total WHERE order_id = @oid", new[] { new NpgsqlParameter("@total", newTotal), new NpgsqlParameter("@oid", orderId) });
                MessageBox.Show("Позиция удалена.");
            };

            btnSaveOrder.Click += (s, ev) =>
            {
                if (cmbClient.SelectedValue == null || cmbManager.SelectedValue == null) { MessageBox.Show("Выберите клиента и менеджера"); return; }
                int clientId = (int)cmbClient.SelectedValue;
                int managerId = (int)cmbManager.SelectedValue;
                string status = cmbStatus.SelectedItem?.ToString() ?? "New";
                string payment = txtPayment.Text.Trim(); if (string.IsNullOrEmpty(payment)) payment = "Наличные";
                DateTime delivery = dtpDelivery.Value;
                if (orderId == 0)
                {
                    string insertOrder = @"
                        INSERT INTO orders (order_date, client_id, manager_id, status, total_amount, payment, delivery_date)
                        VALUES (CURRENT_DATE, @client, @manager, @status, 0, @payment, @delivery)
                        RETURNING order_id";
                    DataTable res = DbConnection.ExecuteQuery(insertOrder, new[] { new NpgsqlParameter("@client", clientId), new NpgsqlParameter("@manager", managerId), new NpgsqlParameter("@status", status), new NpgsqlParameter("@payment", payment), new NpgsqlParameter("@delivery", delivery) });
                    int newOrderId = (int)res.Rows[0]["order_id"];
                    DbConnection.ExecuteNonQuery("INSERT INTO order_status_log (order_id, old_status, new_status, change_date, employee_id) VALUES (@oid, NULL, @status, CURRENT_TIMESTAMP, @emp)", new[] { new NpgsqlParameter("@oid", newOrderId), new NpgsqlParameter("@status", status), new NpgsqlParameter("@emp", UserSession.EmployeeId) });
                    MessageBox.Show($"Заказ №{newOrderId} создан. Теперь нажмите 'Редактировать' для добавления позиций.");
                }
                else
                {
                    string update = @"UPDATE orders SET client_id=@client, manager_id=@manager, status=@status, payment=@payment, delivery_date=@delivery WHERE order_id=@oid";
                    DbConnection.ExecuteNonQuery(update, new[] { new NpgsqlParameter("@client", clientId), new NpgsqlParameter("@manager", managerId), new NpgsqlParameter("@status", status), new NpgsqlParameter("@payment", payment), new NpgsqlParameter("@delivery", delivery), new NpgsqlParameter("@oid", orderId) });
                    MessageBox.Show("Заказ обновлён.");
                }
                dialog.Close();
                LoadOrders();
            };
            btnCancel.Click += (s, ev) => dialog.Close();
            dialog.ShowDialog();
        }

        private void ExportToExcel()
        {
            if (dgvOrders.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excel = Activator.CreateInstance(excelType);
                excel.Visible = true;
                var workbook = excel.Workbooks.Add();
                var sheet = workbook.ActiveSheet;
                for (int i = 0; i < dgvOrders.Columns.Count; i++)
                    sheet.Cells[1, i + 1] = dgvOrders.Columns[i].HeaderText;
                for (int i = 0; i < dgvOrders.Rows.Count; i++)
                    for (int j = 0; j < dgvOrders.Columns.Count; j++)
                        sheet.Cells[i + 2, j + 1] = dgvOrders.Rows[i].Cells[j].Value?.ToString();
                MessageBox.Show("Экспорт в Excel выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Excel: " + ex.Message); }
        }

        private void ExportToWord()
        {
            if (dgvOrders.Rows.Count == 0) return;
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                dynamic word = Activator.CreateInstance(wordType);
                word.Visible = true;
                var doc = word.Documents.Add();
                int rows = dgvOrders.Rows.Count + 1;
                int cols = dgvOrders.Columns.Count;
                var table = doc.Tables.Add(doc.Range(), rows, cols);
                for (int i = 0; i < cols; i++)
                    table.Cell(1, i + 1).Range.Text = dgvOrders.Columns[i].HeaderText;
                for (int i = 0; i < dgvOrders.Rows.Count; i++)
                    for (int j = 0; j < cols; j++)
                        table.Cell(i + 2, j + 1).Range.Text = dgvOrders.Rows[i].Cells[j].Value?.ToString() ?? "";
                MessageBox.Show("Экспорт в Word выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Word: " + ex.Message); }
        }

        private void ExportToCsv()
        {
            if (dgvOrders.Rows.Count == 0) { MessageBox.Show("Нет данных."); return; }
            SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV файлы (*.csv)|*.csv", FileName = $"orders_{DateTime.Now:yyyyMMdd_HHmmss}.csv" };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8))
                {
                    for (int i = 0; i < dgvOrders.Columns.Count; i++)
                        sw.Write(dgvOrders.Columns[i].HeaderText + (i < dgvOrders.Columns.Count - 1 ? ";" : ""));
                    sw.WriteLine();
                    foreach (DataGridViewRow row in dgvOrders.Rows)
                    {
                        for (int i = 0; i < dgvOrders.Columns.Count; i++)
                            sw.Write(row.Cells[i].Value?.ToString() + (i < dgvOrders.Columns.Count - 1 ? ";" : ""));
                        sw.WriteLine();
                    }
                }
                MessageBox.Show("CSV экспорт завершён.");
            }
        }
    }
}