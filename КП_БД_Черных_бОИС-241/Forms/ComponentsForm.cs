using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class ComponentsForm : Form
    {
        private DataGridView dgvComponents;
        private DataTable componentsTable;
        private Button btnAdd, btnEdit, btnDelete, btnRefresh, btnExcel, btnWord;

        public ComponentsForm()
        {
            InitializeComponents();
            LoadData();
        }

        private void InitializeComponents()
        {
            this.Text = "Комплектующие";
            this.Size = new Size(1100, 500);
            this.StartPosition = FormStartPosition.CenterScreen;

            dgvComponents = new DataGridView()
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };
            this.Controls.Add(dgvComponents);

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
            string query = @"
                SELECT c.component_id, c.name, s.name as supplier_name, ct.type_name, 
                       c.retail_price, c.purchase_price, c.stock_quantity,
                       c.supplier_id, c.type_id
                FROM components c
                JOIN suppliers s ON c.supplier_id = s.supplier_id
                JOIN component_types ct ON c.type_id = ct.type_id
                ORDER BY c.component_id";
            componentsTable = DbConnection.ExecuteQuery(query);
            dgvComponents.DataSource = componentsTable;
            if (dgvComponents.Columns.Contains("component_id"))
                dgvComponents.Columns["component_id"].Visible = false;
            if (dgvComponents.Columns.Contains("supplier_id"))
                dgvComponents.Columns["supplier_id"].Visible = false;
            if (dgvComponents.Columns.Contains("type_id"))
                dgvComponents.Columns["type_id"].Visible = false;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            ShowComponentDialog(0);
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgvComponents.CurrentRow == null) return;
            int id = Convert.ToInt32(dgvComponents.CurrentRow.Cells["component_id"].Value);
            ShowComponentDialog(id);
        }

        private void ShowComponentDialog(int id)
        {
         
            DataTable suppliers = DbConnection.ExecuteQuery("SELECT supplier_id, name FROM suppliers");
            DataTable types = DbConnection.ExecuteQuery("SELECT type_id, type_name FROM component_types");

       
            string oldName = "";
            int oldSupplierId = 0, oldTypeId = 0;
            decimal oldRetail = 0, oldPurchase = 0;
            int oldStock = 0;

            if (id != 0)
            {
                DataRow[] rows = componentsTable.Select($"component_id = {id}");
                if (rows.Length > 0)
                {
                    oldName = rows[0]["name"].ToString();
                    oldSupplierId = Convert.ToInt32(rows[0]["supplier_id"]);
                    oldTypeId = Convert.ToInt32(rows[0]["type_id"]);
                    oldRetail = Convert.ToDecimal(rows[0]["retail_price"]);
                    oldPurchase = Convert.ToDecimal(rows[0]["purchase_price"]);
                    oldStock = Convert.ToInt32(rows[0]["stock_quantity"]);
                }
            }

            Form dialog = new Form()
            {
                Text = id == 0 ? "Добавление комплектующей" : "Редактирование комплектующей",
                Size = new Size(450, 380),
                StartPosition = FormStartPosition.CenterParent
            };

            // Название
            Label lblName = new Label() { Text = "Название:", Location = new Point(20, 20), Size = new Size(100, 25) };
            TextBox txtName = new TextBox() { Text = oldName, Location = new Point(130, 20), Width = 250 };

            // Поставщик
            Label lblSupplier = new Label() { Text = "Поставщик:", Location = new Point(20, 60), Size = new Size(100, 25) };
            ComboBox cmbSupplier = new ComboBox()
            {
                DataSource = suppliers,
                DisplayMember = "name",
                ValueMember = "supplier_id",
                Location = new Point(130, 60),
                Width = 250
            };
            if (oldSupplierId != 0) cmbSupplier.SelectedValue = oldSupplierId;

            // Тип
            Label lblType = new Label() { Text = "Тип:", Location = new Point(20, 100), Size = new Size(100, 25) };
            ComboBox cmbType = new ComboBox()
            {
                DataSource = types,
                DisplayMember = "type_name",
                ValueMember = "type_id",
                Location = new Point(130, 100),
                Width = 250
            };
            if (oldTypeId != 0) cmbType.SelectedValue = oldTypeId;

            // Цены
            Label lblRetail = new Label() { Text = "Розничная цена:", Location = new Point(20, 140), Size = new Size(100, 25) };
            NumericUpDown nudRetail = new NumericUpDown() { Location = new Point(130, 140), Width = 150, DecimalPlaces = 2, Minimum = 0, Maximum = 1000000, Value = oldRetail };

            Label lblPurchase = new Label() { Text = "Закупочная цена:", Location = new Point(20, 180), Size = new Size(100, 25) };
            NumericUpDown nudPurchase = new NumericUpDown() { Location = new Point(130, 180), Width = 150, DecimalPlaces = 2, Minimum = 0, Maximum = 1000000, Value = oldPurchase };

            Label lblStock = new Label() { Text = "Остаток:", Location = new Point(20, 220), Size = new Size(100, 25) };
            NumericUpDown nudStock = new NumericUpDown() { Location = new Point(130, 220), Width = 150, Minimum = 0, Maximum = 1000000, Value = oldStock };

            Button btnSave = new Button() { Text = "Сохранить", Location = new Point(130, 270), Size = new Size(100, 30) };
            Button btnCancel = new Button() { Text = "Отмена", Location = new Point(250, 270), Size = new Size(100, 30) };

            btnSave.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(txtName.Text))
                {
                    MessageBox.Show("Введите название");
                    return;
                }
                if (cmbSupplier.SelectedValue == null || cmbType.SelectedValue == null)
                {
                    MessageBox.Show("Выберите поставщика и тип");
                    return;
                }
                int supplierId = (int)cmbSupplier.SelectedValue;
                int typeId = (int)cmbType.SelectedValue;
                decimal retail = nudRetail.Value;
                decimal purchase = nudPurchase.Value;
                int stock = (int)nudStock.Value;

                if (id == 0)
                {
                    string query = @"INSERT INTO components (name, supplier_id, type_id, retail_price, purchase_price, stock_quantity)
                                     VALUES (@name, @supp, @type, @retail, @purchase, @stock)";
                    NpgsqlParameter[] pars = {
                        new NpgsqlParameter("@name", txtName.Text),
                        new NpgsqlParameter("@supp", supplierId),
                        new NpgsqlParameter("@type", typeId),
                        new NpgsqlParameter("@retail", retail),
                        new NpgsqlParameter("@purchase", purchase),
                        new NpgsqlParameter("@stock", stock)
                    };
                    DbConnection.ExecuteNonQuery(query, pars);
                }
                else
                {
                    string query = @"UPDATE components SET name=@name, supplier_id=@supp, type_id=@type,
                                      retail_price=@retail, purchase_price=@purchase, stock_quantity=@stock
                                      WHERE component_id=@id";
                    NpgsqlParameter[] pars = {
                        new NpgsqlParameter("@name", txtName.Text),
                        new NpgsqlParameter("@supp", supplierId),
                        new NpgsqlParameter("@type", typeId),
                        new NpgsqlParameter("@retail", retail),
                        new NpgsqlParameter("@purchase", purchase),
                        new NpgsqlParameter("@stock", stock),
                        new NpgsqlParameter("@id", id)
                    };
                    DbConnection.ExecuteNonQuery(query, pars);
                }
                LoadData();
                dialog.Close();
            };
            btnCancel.Click += (s, ev) => dialog.Close();

            dialog.Controls.Add(lblName);
            dialog.Controls.Add(txtName);
            dialog.Controls.Add(lblSupplier);
            dialog.Controls.Add(cmbSupplier);
            dialog.Controls.Add(lblType);
            dialog.Controls.Add(cmbType);
            dialog.Controls.Add(lblRetail);
            dialog.Controls.Add(nudRetail);
            dialog.Controls.Add(lblPurchase);
            dialog.Controls.Add(nudPurchase);
            dialog.Controls.Add(lblStock);
            dialog.Controls.Add(nudStock);
            dialog.Controls.Add(btnSave);
            dialog.Controls.Add(btnCancel);

            dialog.ShowDialog();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvComponents.CurrentRow == null) return;
            int id = Convert.ToInt32(dgvComponents.CurrentRow.Cells["component_id"].Value);
            if (MessageBox.Show("Удалить комплектующую?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DbConnection.ExecuteNonQuery("DELETE FROM components WHERE component_id = @id", new[] { new NpgsqlParameter("@id", id) });
                LoadData();
            }
        }

        private void ExportToExcel()
        {
            if (dgvComponents.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excel = Activator.CreateInstance(excelType);
                excel.Visible = true;
                var workbook = excel.Workbooks.Add();
                var sheet = workbook.ActiveSheet;
                for (int i = 0; i < dgvComponents.Columns.Count; i++)
                    sheet.Cells[1, i + 1] = dgvComponents.Columns[i].HeaderText;
                for (int i = 0; i < dgvComponents.Rows.Count; i++)
                    for (int j = 0; j < dgvComponents.Columns.Count; j++)
                        sheet.Cells[i + 2, j + 1] = dgvComponents.Rows[i].Cells[j].Value?.ToString();
                MessageBox.Show("Экспорт в Excel выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Excel: " + ex.Message); }
        }

        private void ExportToWord()
        {
            if (dgvComponents.Rows.Count == 0) return;
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                dynamic word = Activator.CreateInstance(wordType);
                word.Visible = true;
                var doc = word.Documents.Add();
                int rows = dgvComponents.Rows.Count;
                int cols = dgvComponents.Columns.Count;
                var table = doc.Tables.Add(doc.Range(), rows, cols);
                for (int i = 0; i < cols; i++)
                    table.Cell(1, i + 1).Range.Text = dgvComponents.Columns[i].HeaderText;
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                        table.Cell(i + 2, j + 1).Range.Text = dgvComponents.Rows[i].Cells[j].Value?.ToString() ?? "";
                MessageBox.Show("Экспорт в Word выполнен.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Word: " + ex.Message); }
        }
    }
}