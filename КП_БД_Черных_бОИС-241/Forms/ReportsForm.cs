using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Npgsql;
using КП_БД_Черных.Database;

namespace КП_БД_Черных.Forms
{
    public partial class ReportsForm : Form
    {
        private TabControl tabControl;

      
        private DataGridView dgvSales;
        private DateTimePicker dtpSalesFrom;
        private DateTimePicker dtpSalesTo;

        private DataGridView dgvTopClients;

        private DataGridView dgvTopConfigs;

        private DataGridView dgvStock;
        private ComboBox cmbStockType;

        private DataGridView dgvManagers;

        private DataGridView dgvInventory;
        private ComboBox cmbInvType;

        private DataGridView dgvPriceHistory;

        public ReportsForm()
        {
            InitializeComponents();
            LoadAllData();
        }

        private void InitializeComponents()
        {
            this.Text = "Модуль статистики и аналитики";
            this.Size = new Size(1100, 650);
            this.StartPosition = FormStartPosition.CenterScreen;

            tabControl = new TabControl() { Dock = DockStyle.Fill };


            TabPage tpSales = new TabPage("Продажи по периодам");
            Panel pnlSalesTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            dtpSalesFrom = new DateTimePicker() { Format = DateTimePickerFormat.Short, Location = new Point(30, 8), Width = 120 };
            dtpSalesTo = new DateTimePicker() { Format = DateTimePickerFormat.Short, Location = new Point(170, 8), Width = 120 };
            dtpSalesFrom.Value = new DateTime(DateTime.Now.Year, 1, 1);
            dtpSalesTo.Value = DateTime.Now;
            Button btnSalesRefresh = new Button() { Text = "Обновить", Location = new Point(350, 6), Size = new Size(90, 25) };
            btnSalesRefresh.Click += (s, e) => LoadSales();
            pnlSalesTop.Controls.Add(new Label() { Text = "С:", Location = new Point(2, 12), Size = new Size(20, 15) });
            pnlSalesTop.Controls.Add(dtpSalesFrom);
            pnlSalesTop.Controls.Add(new Label() { Text = "По:", Location = new Point(118, 12), Size = new Size(25, 15) });
            pnlSalesTop.Controls.Add(dtpSalesTo);
            pnlSalesTop.Controls.Add(btnSalesRefresh);
            dgvSales = CreateGrid();
            tpSales.Controls.Add(dgvSales);
            tpSales.Controls.Add(pnlSalesTop);
            tabControl.TabPages.Add(tpSales);

      
            TabPage tpClients = new TabPage("Топ клиентов");
            Panel pnlClientsTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            Button btnClientsRefresh = new Button() { Text = "Обновить", Location = new Point(10, 6), Size = new Size(90, 25) };
            btnClientsRefresh.Click += (s, e) => LoadTopClients();
            pnlClientsTop.Controls.Add(btnClientsRefresh);
            dgvTopClients = CreateGrid();
            tpClients.Controls.Add(dgvTopClients);
            tpClients.Controls.Add(pnlClientsTop);
            tabControl.TabPages.Add(tpClients);

   
            TabPage tpConfigs = new TabPage("Популярные сборки");
            Panel pnlConfigsTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            Button btnConfigsRefresh = new Button() { Text = "Обновить", Location = new Point(10, 6), Size = new Size(90, 25) };
            btnConfigsRefresh.Click += (s, e) => LoadTopConfigs();
            pnlConfigsTop.Controls.Add(btnConfigsRefresh);
            dgvTopConfigs = CreateGrid();
            tpConfigs.Controls.Add(dgvTopConfigs);
            tpConfigs.Controls.Add(pnlConfigsTop);
            tabControl.TabPages.Add(tpConfigs);

            TabPage tpStock = new TabPage("Склад");
            Panel pnlStockTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            cmbStockType = new ComboBox() { Location = new Point(10, 6), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            Button btnStockRefresh = new Button() { Text = "Обновить", Location = new Point(200, 6), Size = new Size(90, 25) };
            btnStockRefresh.Click += (s, e) => LoadStock();
            pnlStockTop.Controls.Add(new Label() { Text = "Тип:", Location = new Point(8, 10), Size = new Size(30, 15) });
            pnlStockTop.Controls.Add(cmbStockType);
            pnlStockTop.Controls.Add(btnStockRefresh);
            dgvStock = CreateGrid();
            dgvStock.CellFormatting += DgvStock_CellFormatting;
            tpStock.Controls.Add(dgvStock);
            tpStock.Controls.Add(pnlStockTop);
            tabControl.TabPages.Add(tpStock);


            TabPage tpManagers = new TabPage("Эффективность менеджеров");
            Panel pnlManagersTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            Button btnManagersRefresh = new Button() { Text = "Обновить", Location = new Point(10, 6), Size = new Size(90, 25) };
            btnManagersRefresh.Click += (s, e) => LoadManagers();
            pnlManagersTop.Controls.Add(btnManagersRefresh);
            dgvManagers = CreateGrid();
            tpManagers.Controls.Add(dgvManagers);
            tpManagers.Controls.Add(pnlManagersTop);
            tabControl.TabPages.Add(tpManagers);

 
            TabPage tpInv = new TabPage("Движение по складу");
            Panel pnlInvTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            cmbInvType = new ComboBox() { Location = new Point(10, 6), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInvType.Items.AddRange(new object[] { "Все", "purchase", "sale", "return" });
            cmbInvType.SelectedIndex = 0;
            Button btnInvRefresh = new Button() { Text = "Обновить", Location = new Point(160, 6), Size = new Size(90, 25) };
            btnInvRefresh.Click += (s, e) => LoadInventory();
            pnlInvTop.Controls.Add(new Label() { Text = "Тип:", Location = new Point(8, 10), Size = new Size(30, 15) });
            pnlInvTop.Controls.Add(cmbInvType);
            pnlInvTop.Controls.Add(btnInvRefresh);
            dgvInventory = CreateGrid();
            tpInv.Controls.Add(dgvInventory);
            tpInv.Controls.Add(pnlInvTop);
            tabControl.TabPages.Add(tpInv);

 
            TabPage tpPrices = new TabPage("История изменения цен");
            Panel pnlPricesTop = new Panel() { Dock = DockStyle.Top, Height = 40 };
            Button btnPricesRefresh = new Button() { Text = "Обновить", Location = new Point(10, 6), Size = new Size(90, 25) };
            btnPricesRefresh.Click += (s, e) => LoadPriceHistory();
            pnlPricesTop.Controls.Add(btnPricesRefresh);
            dgvPriceHistory = CreateGrid();
            tpPrices.Controls.Add(dgvPriceHistory);
            tpPrices.Controls.Add(pnlPricesTop);
            tabControl.TabPages.Add(tpPrices);

            Panel pnlBottom = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            Button btnExcel = new Button() { Text = "Excel", Location = new Point(10, 10), Size = new Size(90, 30) };
            Button btnWord = new Button() { Text = "Word", Location = new Point(110, 10), Size = new Size(90, 30) };
            Button btnCsv = new Button() { Text = "CSV", Location = new Point(210, 10), Size = new Size(90, 30) };
            Button btnRefreshAll = new Button() { Text = "Обновить всё", Location = new Point(320, 10), Size = new Size(110, 30) };

            btnExcel.Click += (s, e) => ExportToExcel(GetActiveGrid());
            btnWord.Click += (s, e) => ExportToWord(GetActiveGrid());
            btnCsv.Click += (s, e) => ExportToCsv(GetActiveGrid());
            btnRefreshAll.Click += (s, e) => LoadAllData();

            pnlBottom.Controls.Add(btnExcel);
            pnlBottom.Controls.Add(btnWord);
            pnlBottom.Controls.Add(btnCsv);
            pnlBottom.Controls.Add(btnRefreshAll);

            this.Controls.Add(tabControl);
            this.Controls.Add(pnlBottom);
        }

        private DataGridView CreateGrid()
        {
            return new DataGridView()
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                BackgroundColor = Color.White
            };
        }

        private DataGridView GetActiveGrid()
        {
            switch (tabControl.SelectedIndex)
            {
                case 0: return dgvSales;
                case 1: return dgvTopClients;
                case 2: return dgvTopConfigs;
                case 3: return dgvStock;
                case 4: return dgvManagers;
                case 5: return dgvInventory;
                case 6: return dgvPriceHistory;
                default: return dgvSales;
            }
        }

        private void LoadAllData()
        {
            LoadSales();
            LoadTopClients();
            LoadTopConfigs();
            LoadStockTypes(); 
            LoadStock();
            LoadManagers();
            LoadInventory();
            LoadPriceHistory();
        }

        private void LoadSales()
        {
            string query = @"
                SELECT period, total_revenue, orders_count, avg_check, rating
                FROM sales_statistics
                WHERE period BETWEEN @d1 AND @d2
                ORDER BY period";
            NpgsqlParameter[] parameters = {
                new NpgsqlParameter("@d1", dtpSalesFrom.Value.Date),
                new NpgsqlParameter("@d2", dtpSalesTo.Value.Date)
            };
            dgvSales.DataSource = DbConnection.ExecuteQuery(query, parameters);
            SetColumnHeaders(dgvSales, new[] {
                ("period", "Период"),
                ("total_revenue", "Выручка"),
                ("orders_count", "Заказов"),
                ("avg_check", "Средний чек"),
                ("rating", "Рейтинг")
            });
        }


        private void LoadTopClients()
        {
            string query = @"
                SELECT 
                    c.surname || ' ' || c.name || ' ' || COALESCE(c.patronymic, '') AS client,
                    COUNT(o.order_id) AS orders_count,
                    COALESCE(SUM(o.total_amount), 0) AS total_spent
                FROM clients c
                LEFT JOIN orders o ON c.client_id = o.client_id
                GROUP BY c.client_id, c.surname, c.name, c.patronymic
                ORDER BY total_spent DESC";
            dgvTopClients.DataSource = DbConnection.ExecuteQuery(query);
            SetColumnHeaders(dgvTopClients, new[] {
                ("client", "Клиент"),
                ("orders_count", "Заказов"),
                ("total_spent", "Сумма заказов")
            });
        }

       
        private void LoadTopConfigs()
        {
            string query = @"
                SELECT 
                    cfg.name AS configuration,
                    COALESCE(SUM(oi.quantity), 0) AS sold_count,
                    COALESCE(SUM(oi.quantity * oi.unit_price), 0) AS total_revenue
                FROM configurations cfg
                LEFT JOIN order_items oi ON cfg.configuration_id = oi.configuration_id
                LEFT JOIN orders o ON oi.order_id = o.order_id
                GROUP BY cfg.configuration_id, cfg.name
                ORDER BY total_revenue DESC";
            dgvTopConfigs.DataSource = DbConnection.ExecuteQuery(query);
            SetColumnHeaders(dgvTopConfigs, new[] {
                ("configuration", "Конфигурация"),
                ("sold_count", "Продано шт."),
                ("total_revenue", "Выручка")
            });
        }

 
        private void LoadStockTypes()
        {
            DataTable dt = DbConnection.ExecuteQuery("SELECT type_id, type_name FROM component_types ORDER BY type_name");
            DataRow allRow = dt.NewRow();
            allRow["type_id"] = 0;
            allRow["type_name"] = "Все";
            dt.Rows.InsertAt(allRow, 0);
            cmbStockType.DataSource = dt;
            cmbStockType.DisplayMember = "type_name";
            cmbStockType.ValueMember = "type_id";
        }

        private void LoadStock()
        {
            int typeId = cmbStockType.SelectedValue == null ? 0 : Convert.ToInt32(cmbStockType.SelectedValue);
            string query = @"
                SELECT 
                    c.name AS component,
                    ct.type_name AS type,
                    s.name AS supplier,
                    c.stock_quantity,
                    c.retail_price,
                    c.purchase_price,
                    (c.retail_price - c.purchase_price) AS margin
                FROM components c
                JOIN component_types ct ON c.type_id = ct.type_id
                JOIN suppliers s ON c.supplier_id = s.supplier_id
                WHERE (@typeId = 0 OR c.type_id = @typeId)
                ORDER BY c.stock_quantity ASC";
            NpgsqlParameter[] parameters = {
                new NpgsqlParameter("@typeId", typeId)
            };
            dgvStock.DataSource = DbConnection.ExecuteQuery(query, parameters);
            SetColumnHeaders(dgvStock, new[] {
                ("component", "Комплектующее"),
                ("type", "Тип"),
                ("supplier", "Поставщик"),
                ("stock_quantity", "Остаток"),
                ("retail_price", "Розн. цена"),
                ("purchase_price", "Закуп. цена"),
                ("margin", "Наценка")
            });
        }

        private void DgvStock_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvStock.Columns[e.ColumnIndex].Name == "stock_quantity" && e.Value != null)
            {
                int qty = Convert.ToInt32(e.Value);
                if (qty <= 3)
                    dgvStock.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightCoral;
                else if (qty <= 5)
                    dgvStock.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                else
                    dgvStock.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
            }
        }


        private void LoadManagers()
        {
            string query = @"
                SELECT 
                    e.full_name AS manager,
                    COUNT(o.order_id) AS orders_count,
                    COALESCE(SUM(o.total_amount), 0) AS total_revenue,
                    ROUND(COALESCE(AVG(o.total_amount), 0), 2) AS avg_check
                FROM employees e
                LEFT JOIN orders o ON e.employee_id = o.manager_id
                WHERE e.position = 'Manager'
                GROUP BY e.employee_id, e.full_name
                ORDER BY total_revenue DESC";
            dgvManagers.DataSource = DbConnection.ExecuteQuery(query);
            SetColumnHeaders(dgvManagers, new[] {
                ("manager", "Менеджер"),
                ("orders_count", "Заказов"),
                ("total_revenue", "Выручка"),
                ("avg_check", "Средний чек")
            });
        }


        private void LoadInventory()
        {
            string typeFilter = cmbInvType.SelectedItem?.ToString() ?? "Все";
            string query = @"
                SELECT 
                    c.name AS component,
                    im.movement_type,
                    im.quantity,
                    im.movement_date,
                    im.reference_id
                FROM inventory_movements im
                JOIN components c ON im.component_id = c.component_id
                WHERE (@type = 'Все' OR im.movement_type = @type)
                ORDER BY im.movement_date DESC";
            NpgsqlParameter[] parameters = {
                new NpgsqlParameter("@type", typeFilter)
            };
            dgvInventory.DataSource = DbConnection.ExecuteQuery(query, parameters);
            SetColumnHeaders(dgvInventory, new[] {
                ("component", "Комплектующее"),
                ("movement_type", "Тип движения"),
                ("quantity", "Количество"),
                ("movement_date", "Дата"),
                ("reference_id", "ID ссылки")
            });
        }

     
        private void LoadPriceHistory()
        {
            string query = @"
                SELECT 
                    c.name AS component,
                    cph.old_price,
                    cph.new_price,
                    cph.change_date
                FROM component_price_history cph
                JOIN components c ON cph.component_id = c.component_id
                ORDER BY cph.change_date DESC";
            dgvPriceHistory.DataSource = DbConnection.ExecuteQuery(query);
            SetColumnHeaders(dgvPriceHistory, new[] {
                ("component", "Комплектующее"),
                ("old_price", "Старая цена"),
                ("new_price", "Новая цена"),
                ("change_date", "Дата изменения")
            });
        }

        private void SetColumnHeaders(DataGridView dgv, (string name, string header)[] headers)
        {
            foreach (var h in headers)
            {
                if (dgv.Columns.Contains(h.name))
                    dgv.Columns[h.name].HeaderText = h.header;
            }
        }


        private void ExportToExcel(DataGridView dgv)
        {
            if (dgv == null || dgv.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic excel = Activator.CreateInstance(excelType);
                excel.Visible = true;
                var workbook = excel.Workbooks.Add();
                var sheet = workbook.ActiveSheet;
                for (int i = 0; i < dgv.Columns.Count; i++)
                    sheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
                for (int i = 0; i < dgv.Rows.Count; i++)
                    for (int j = 0; j < dgv.Columns.Count; j++)
                        sheet.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value?.ToString();
                MessageBox.Show("Экспорт в Excel выполнен.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка Excel: " + ex.Message);
            }
        }

        private void ExportToWord(DataGridView dgv)
        {
            if (dgv == null || dgv.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                dynamic word = Activator.CreateInstance(wordType);
                word.Visible = true;
                var doc = word.Documents.Add();
                int rows = dgv.Rows.Count + 1;
                int cols = dgv.Columns.Count;
                var table = doc.Tables.Add(doc.Range(), rows, cols);
                for (int i = 0; i < cols; i++)
                    table.Cell(1, i + 1).Range.Text = dgv.Columns[i].HeaderText;
                for (int i = 0; i < dgv.Rows.Count; i++)
                    for (int j = 0; j < cols; j++)
                        table.Cell(i + 2, j + 1).Range.Text = dgv.Rows[i].Cells[j].Value?.ToString() ?? "";
                MessageBox.Show("Экспорт в Word выполнен.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка Word: " + ex.Message);
            }
        }

        private void ExportToCsv(DataGridView dgv)
        {
            if (dgv == null || dgv.Rows.Count == 0) { MessageBox.Show("Нет данных для экспорта."); return; }
            SaveFileDialog sfd = new SaveFileDialog()
            {
                Filter = "CSV файлы (*.csv)|*.csv",
                FileName = $"report_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.UTF8))
                {
                    for (int i = 0; i < dgv.Columns.Count; i++)
                        sw.Write(dgv.Columns[i].HeaderText + (i < dgv.Columns.Count - 1 ? ";" : ""));
                    sw.WriteLine();
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        for (int i = 0; i < dgv.Columns.Count; i++)
                            sw.Write(row.Cells[i].Value?.ToString() + (i < dgv.Columns.Count - 1 ? ";" : ""));
                        sw.WriteLine();
                    }
                }
                MessageBox.Show("CSV экспорт завершён.");
            }
        }
    }
}
