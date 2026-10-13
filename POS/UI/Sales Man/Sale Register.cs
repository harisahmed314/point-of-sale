using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using DGVPrinterHelper;
using POS.BLL;
using POS.DAL;

namespace POS.UI.Sales_Man
{
    public partial class Sale_Register : Form
    {
        private static Sale_Register _instance;

        public static Sale_Register instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Sale_Register();

                return _instance;
            }
        }

        // =========================================================
        // OBJECTS
        // =========================================================

        private DataTable table = new DataTable();

        Manage_Productsbll u = new Manage_Productsbll();
        Manage_Productsdal dal = new Manage_Productsdal();

        invoicebll inbll = new invoicebll();
        invoicedal indal = new invoicedal();

        userdal user = new userdal();

        product_detailsdal prdal = new product_detailsdal();
        product_detailsbll prbll = new product_detailsbll();

        // =========================================================
        // SELECTED PRODUCT VARIABLES
        // =========================================================

        private int selectedProductId = 0;

        private decimal selectedStockQuantity = 0;
        private decimal selectedPrice = 0;
        private string selectedType = "";
        private string selectedCode = "";

        // =========================================================
        // CALCULATION VARIABLES
        // =========================================================

        private bool updatingQuantity = false;
        private bool updatingDiscount = false;
        private bool updatingPaid = false;

        private decimal totalPayable = 0;
        private decimal paidAmount = 0;

        // =========================================================
        // CONSTRUCTOR
        // =========================================================

        public Sale_Register()
        {
            InitializeComponent();
        }

        // =========================================================
        // FORM LOAD
        // =========================================================

        private void Sale_Register_Load(object sender, EventArgs e)
        {
            try
            {
                LoadStock();
                SetupInvoiceTable();

                SetDefaultValues();

                ToGetInvoiceID();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error while loading Sales Register.\n\n" + ex.Message,
                    "Load Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // LOAD STOCK
        // =========================================================

        private void LoadStock()
        {
            DataTable dt = dal.select();

            if (dt != null)
            {
                dgvstock.DataSource = dt;
            }
        }

        // =========================================================
        // SETUP INVOICE TABLE
        // =========================================================

        private void SetupInvoiceTable()
        {
            table = new DataTable();

            // IMPORTANT:
            // ProductId is stored separately.
            // This fixes the problem where "blk" was being
            // converted to integer.

            table.Columns.Add("ProductId", typeof(int));
            table.Columns.Add("Item Name", typeof(string));
            table.Columns.Add("Code", typeof(string));
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Price", typeof(decimal));
            table.Columns.Add("Quantity", typeof(decimal));
            table.Columns.Add("Total", typeof(decimal));

            dgvinvoice.AutoGenerateColumns = true;
            dgvinvoice.DataSource = table;

            ConfigureInvoiceGrid();
        }

        // =========================================================
        // CONFIGURE INVOICE GRID
        // =========================================================

        private void ConfigureInvoiceGrid()
        {
            if (dgvinvoice.Columns.Count == 0)
                return;

            // Hide ProductId from user
            if (dgvinvoice.Columns.Contains("ProductId"))
            {
                dgvinvoice.Columns["ProductId"].Visible = false;
            }

            if (dgvinvoice.Columns.Contains("Item Name"))
                dgvinvoice.Columns["Item Name"].Width = 120;

            if (dgvinvoice.Columns.Contains("Code"))
                dgvinvoice.Columns["Code"].Width = 70;

            if (dgvinvoice.Columns.Contains("Type"))
                dgvinvoice.Columns["Type"].Width = 80;

            if (dgvinvoice.Columns.Contains("Price"))
                dgvinvoice.Columns["Price"].Width = 80;

            if (dgvinvoice.Columns.Contains("Quantity"))
                dgvinvoice.Columns["Quantity"].Width = 70;

            if (dgvinvoice.Columns.Contains("Total"))
                dgvinvoice.Columns["Total"].Width = 90;

            if (dgvinvoice.Columns.Contains("Price"))
                dgvinvoice.Columns["Price"].DefaultCellStyle.Format = "0.##";

            if (dgvinvoice.Columns.Contains("Quantity"))
                dgvinvoice.Columns["Quantity"].DefaultCellStyle.Format = "0.##";

            if (dgvinvoice.Columns.Contains("Total"))
                dgvinvoice.Columns["Total"].DefaultCellStyle.Format = "0.##";
        }

        // =========================================================
        // DEFAULT VALUES
        // =========================================================

        private void SetDefaultValues()
        {
            txttotal.Text = "0";
            txtdiscount.Text = "0";
            txtsubtotal.Text = "0";
            txtpaidamount.Text = "0";
            txtdueamount.Text = "0";
            txtchangeamount.Text = "0";

            txtitemname.Text = "";
            txtprice.Text = "";
            txtquantity.Text = "";
            txttotal1.Text = "";

            cmbcustomer.Text = "";
            txtcomment.Text = "";

            selectedProductId = 0;
            selectedStockQuantity = 0;
            selectedPrice = 0;
            selectedType = "";
            selectedCode = "";

            totalPayable = 0;
            paidAmount = 0;
        }

        // =========================================================
        // CLEAR SALE
        // =========================================================

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (table != null)
                {
                    table.Clear();
                }

                SetDefaultValues();

                ClearErrors();

                ToGetInvoiceID();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to clear invoice.\n\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // STOCK GRID PRODUCT SELECTION
        // =========================================================

        private void dgvstock_RowHeaderMouseClick(
            object sender,
            DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0)
                    return;

                DataGridViewRow row = dgvstock.Rows[e.RowIndex];

                // -------------------------------------------------
                // Your existing database column indexes:
                //
                // 1 = Product Name
                // 2 = Code
                // 6 = Selling Price
                // 7 = Type
                // 8 = Quantity
                //
                // Product ID is normally column 0.
                // -------------------------------------------------

                if (row.Cells.Count <= 8)
                {
                    MessageBox.Show(
                        "Stock grid does not contain the expected columns.",
                        "Product Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                // PRODUCT ID
                int productId;

                if (!int.TryParse(
                    Convert.ToString(row.Cells[0].Value),
                    out productId))
                {
                    MessageBox.Show(
                        "Invalid Product ID.",
                        "Product Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                selectedProductId = productId;

                // PRODUCT NAME
                txtitemname.Text =
                    Convert.ToString(row.Cells[1].Value);

                // CODE
                selectedCode =
                    Convert.ToString(row.Cells[2].Value);

                // PRICE
                decimal price;

                if (!TryParseDecimal(
                    Convert.ToString(row.Cells[6].Value),
                    out price))
                {
                    MessageBox.Show(
                        "This product has an invalid selling price.",
                        "Invalid Price",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                selectedPrice = price;

                txtprice.Text =
                    selectedPrice.ToString("0.##");

                // TYPE
                selectedType =
                    Convert.ToString(row.Cells[7].Value);

                // STOCK QUANTITY
                decimal stockQuantity;

                if (!TryParseDecimal(
                    Convert.ToString(row.Cells[8].Value),
                    out stockQuantity))
                {
                    MessageBox.Show(
                        "Invalid stock quantity.",
                        "Stock Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                selectedStockQuantity = stockQuantity;

                if (selectedStockQuantity <= 0)
                {
                    MessageBox.Show(
                        "Product quantity is zero. Please update stock.",
                        "Out of Stock",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                txtquantity.Text = "1";

                CalculateItemTotal();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to select product.\n\n" + ex.Message,
                    "Product Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // SEARCH STOCK
        // =========================================================

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string keyword = txtsearch.Text.Trim();

                DataTable dt;

                if (string.IsNullOrWhiteSpace(keyword))
                {
                    dt = dal.select();
                }
                else
                {
                    dt = dal.Search(keyword);
                }

                dgvstock.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Search error.\n\n" + ex.Message,
                    "Search Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // QUANTITY CHANGED
        // =========================================================

        private void txtquantity_TextChanged(object sender, EventArgs e)
        {
            if (updatingQuantity)
                return;

            CalculateItemTotal();
        }

        // =========================================================
        // CALCULATE ITEM TOTAL
        // =========================================================

        private void CalculateItemTotal()
        {
            if (updatingQuantity)
                return;

            try
            {
                updatingQuantity = true;

                decimal quantity;

                if (!TryParseDecimal(txtquantity.Text, out quantity))
                {
                    txttotal1.Text = "0";
                    return;
                }

                decimal price;

                if (!TryParseDecimal(txtprice.Text, out price))
                {
                    txttotal1.Text = "0";
                    return;
                }

                if (quantity <= 0)
                {
                    txttotal1.Text = "0";
                    return;
                }

                decimal itemTotal = quantity * price;

                txttotal1.Text =
                    itemTotal.ToString("0.##");
            }
            finally
            {
                updatingQuantity = false;
            }
        }

        // =========================================================
        // ADD PRODUCT
        // =========================================================

        private void btnadd_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateProductForAdd())
                    return;

                decimal quantity;
                decimal price;

                if (!TryParseDecimal(
                    txtquantity.Text,
                    out quantity))
                {
                    MessageBox.Show(
                        "Please enter a valid quantity.",
                        "Invalid Quantity",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                if (!TryParseDecimal(
                    txtprice.Text,
                    out price))
                {
                    MessageBox.Show(
                        "Please enter a valid price.",
                        "Invalid Price",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                // ---------------------------------------------
                // CHECK STOCK
                // ---------------------------------------------

                if (quantity > selectedStockQuantity)
                {
                    MessageBox.Show(
                        "Only " +
                        selectedStockQuantity.ToString("0.##") +
                        " item(s) are available in stock.",
                        "Insufficient Stock",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                    return;
                }

                decimal itemTotal = quantity * price;

                // ---------------------------------------------
                // CHECK IF SAME PRODUCT ALREADY EXISTS
                // ---------------------------------------------

                DataRow existingRow = null;

                foreach (DataRow row in table.Rows)
                {
                    int existingProductId =
                        Convert.ToInt32(row["ProductId"]);

                    if (existingProductId == selectedProductId)
                    {
                        existingRow = row;
                        break;
                    }
                }

                if (existingRow != null)
                {
                    decimal oldQuantity =
                        Convert.ToDecimal(existingRow["Quantity"]);

                    decimal newQuantity =
                        oldQuantity + quantity;

                    if (newQuantity > selectedStockQuantity)
                    {
                        MessageBox.Show(
                            "Total quantity of this product cannot exceed available stock.",
                            "Insufficient Stock",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    existingRow["Quantity"] = newQuantity;
                    existingRow["Price"] = price;
                    existingRow["Total"] =
                        newQuantity * price;
                }
                else
                {
                    // -----------------------------------------
                    // ADD NEW PRODUCT
                    // -----------------------------------------

                    table.Rows.Add(
                        selectedProductId,
                        txtitemname.Text.Trim(),
                        selectedCode,
                        selectedType,
                        price,
                        quantity,
                        itemTotal
                    );
                }

                dgvinvoice.DataSource = null;
                dgvinvoice.DataSource = table;

                ConfigureInvoiceGrid();

                RecalculateTotals();

                ClearSelectedProduct();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to add product.\n\n" + ex.Message,
                    "Add Product Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // VALIDATE PRODUCT
        // =========================================================

        private bool ValidateProductForAdd()
        {
            ClearErrors();

            if (selectedProductId <= 0)
            {
                MessageBox.Show(
                    "Please select a product from stock.",
                    "Product Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            if (string.IsNullOrWhiteSpace(txtitemname.Text))
            {
                errorProvider1.SetError(
                    txtitemname,
                    "Product name required."
                );

                return false;
            }

            decimal price;

            if (!TryParseDecimal(
                txtprice.Text,
                out price) ||
                price <= 0)
            {
                errorProvider1.SetError(
                    txtprice,
                    "Valid product price required."
                );

                return false;
            }

            decimal quantity;

            if (!TryParseDecimal(
                txtquantity.Text,
                out quantity) ||
                quantity <= 0)
            {
                errorProvider1.SetError(
                    txtquantity,
                    "Quantity must be greater than zero."
                );

                return false;
            }

            return true;
        }

        // =========================================================
        // CLEAR SELECTED PRODUCT
        // =========================================================

        private void ClearSelectedProduct()
        {
            txtitemname.Text = "";
            txtprice.Text = "";
            txtquantity.Text = "";
            txttotal1.Text = "";

            selectedProductId = 0;
            selectedStockQuantity = 0;
            selectedPrice = 0;
            selectedType = "";
            selectedCode = "";
        }

        private void clear()
        {
            ClearSelectedProduct();
        }

        // =========================================================
        // RECALCULATE EVERYTHING
        // =========================================================

        private void RecalculateTotals()
        {
            try
            {
                decimal total = 0;

                // ---------------------------------------------
                // SUM ALL PRODUCTS
                // ---------------------------------------------

                foreach (DataRow row in table.Rows)
                {
                    decimal rowTotal;

                    if (TryParseDecimal(
                        Convert.ToString(row["Total"]),
                        out rowTotal))
                    {
                        total += rowTotal;
                    }
                }

                txttotal.Text =
                    total.ToString("0.##");

                // ---------------------------------------------
                // DISCOUNT
                // ---------------------------------------------

                decimal discount;

                if (!TryParseDecimal(
                    txtdiscount.Text,
                    out discount))
                {
                    discount = 0;
                }

                if (discount < 0)
                    discount = 0;

                if (discount > 100)
                    discount = 100;

                txtdiscount.Text =
                    discount.ToString("0.##");

                // ---------------------------------------------
                // PAYABLE AFTER DISCOUNT
                // ---------------------------------------------

                decimal discountAmount =
                    total * discount / 100;

                decimal subtotal =
                    total - discountAmount;

                if (subtotal < 0)
                    subtotal = 0;

                txtsubtotal.Text =
                    subtotal.ToString("0.##");

                totalPayable = subtotal;

                // ---------------------------------------------
                // PAID
                // ---------------------------------------------

                decimal paid;

                if (!TryParseDecimal(
                    txtpaidamount.Text,
                    out paid))
                {
                    paid = 0;
                }

                if (paid < 0)
                    paid = 0;

                txtpaidamount.Text =
                    paid.ToString("0.##");

                paidAmount = paid;

                // ---------------------------------------------
                // DUE / CHANGE
                // ---------------------------------------------

                if (paid < subtotal)
                {
                    decimal due =
                        subtotal - paid;

                    txtdueamount.Text =
                        due.ToString("0.##");

                    txtchangeamount.Text = "0";
                }
                else
                {
                    decimal change =
                        paid - subtotal;

                    txtdueamount.Text = "0";

                    txtchangeamount.Text =
                        change.ToString("0.##");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to calculate totals.\n\n" +
                    ex.Message,
                    "Calculation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // DISCOUNT CHANGED
        // =========================================================

        private void txtdiscount_TextChanged(object sender, EventArgs e)
        {
            if (updatingDiscount)
                return;

            try
            {
                updatingDiscount = true;

                decimal discount;

                if (!TryParseDecimal(
                    txtdiscount.Text,
                    out discount))
                {
                    discount = 0;
                }

                if (discount < 0)
                    discount = 0;

                if (discount > 100)
                    discount = 100;

                txtdiscount.Text =
                    discount.ToString("0.##");

                RecalculateTotals();
            }
            finally
            {
                updatingDiscount = false;
            }
        }

        // =========================================================
        // PAID AMOUNT CHANGED
        // =========================================================

        private void txtpaidamount_TextChanged(object sender, EventArgs e)
        {
            if (updatingPaid)
                return;

            try
            {
                updatingPaid = true;

                RecalculateTotals();
            }
            finally
            {
                updatingPaid = false;
            }
        }

        // =========================================================
        // DELETE INVOICE ROW
        // =========================================================

        private void dgvinvoice_RowHeaderMouseDoubleClick(
            object sender,
            DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0)
                    return;

                if (e.RowIndex >= table.Rows.Count)
                    return;

                DialogResult result = MessageBox.Show(
                    "Remove this product from the invoice?",
                    "Remove Product",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result != DialogResult.Yes)
                    return;

                table.Rows.RemoveAt(e.RowIndex);

                RecalculateTotals();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to remove product.\n\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // CLEAR BUTTON
        // =========================================================

        private void button5_Click(object sender, EventArgs e)
        {
            ClearSelectedProduct();
        }

        // =========================================================
        // MAIN SAVE / COMPLETE SALE
        // =========================================================

        private void btncomplete_Click(object sender, EventArgs e)
        {
            SaveSale(true);
        }

        // =========================================================
        // SAVE WITHOUT PRINT
        // =========================================================

        private void btnsave_Click(object sender, EventArgs e)
        {
            SaveSale(false);
        }

        // =========================================================
        // SAVE SALE
        // =========================================================

        private void SaveSale(bool printInvoice)
        {
            try
            {
                // ---------------------------------------------
                // VALIDATE
                // ---------------------------------------------

                if (!Validation1())
                    return;

                // ---------------------------------------------
                // RECALCULATE BEFORE SAVING
                // ---------------------------------------------

                RecalculateTotals();

                decimal subtotal;
                decimal paid;
                decimal discount;
                decimal due;
                decimal change;

                if (!TryParseDecimal(
                    txtsubtotal.Text,
                    out subtotal))
                {
                    MessageBox.Show(
                        "Invalid total payable.",
                        "Sale Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                if (!TryParseDecimal(
                    txtpaidamount.Text,
                    out paid))
                {
                    paid = 0;
                }

                if (!TryParseDecimal(
                    txtdiscount.Text,
                    out discount))
                {
                    discount = 0;
                }

                if (!TryParseDecimal(
                    txtdueamount.Text,
                    out due))
                {
                    due = 0;
                }

                if (!TryParseDecimal(
                    txtchangeamount.Text,
                    out change))
                {
                    change = 0;
                }

                // ---------------------------------------------
                // GET LOGGED USER
                // ---------------------------------------------

                string loggedUser = Login.loggedIn;

                if (string.IsNullOrWhiteSpace(loggedUser))
                {
                    MessageBox.Show(
                        "No logged-in user was found.",
                        "Login Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                usersbll usr =
                    user.GetIDFromUsername(loggedUser);

                if (usr == null)
                {
                    MessageBox.Show(
                        "Unable to find logged-in user.",
                        "User Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                // ---------------------------------------------
                // INVOICE NUMBER
                // ---------------------------------------------

                int invoiceNumber;

                if (!int.TryParse(
                    lblinvoiceno.Text,
                    out invoiceNumber))
                {
                    MessageBox.Show(
                        "Invalid invoice number.",
                        "Invoice Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                // ---------------------------------------------
                // PRE-CHECK ALL PRODUCTS
                // ---------------------------------------------

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int productId;

                    if (!int.TryParse(
                        Convert.ToString(table.Rows[i]["ProductId"]),
                        out productId))
                    {
                        MessageBox.Show(
                            "Invalid product ID for product: " +
                            table.Rows[i]["Item Name"],
                            "Product Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        return;
                    }

                    decimal saleQuantity;

                    if (!TryParseDecimal(
                        Convert.ToString(table.Rows[i]["Quantity"]),
                        out saleQuantity) ||
                        saleQuantity <= 0)
                    {
                        MessageBox.Show(
                            "Invalid quantity for product: " +
                            table.Rows[i]["Item Name"],
                            "Quantity Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        return;
                    }
                }

                // ---------------------------------------------
                // CREATE INVOICE BLL
                // ---------------------------------------------

                inbll.inv_no = invoiceNumber;

                inbll.customer_name =
                    cmbcustomer.Text.Trim();

                inbll.total_payable =
                    subtotal.ToString(
                        "0.##",
                        CultureInfo.InvariantCulture);

                inbll.paid_amount =
                    paid.ToString(
                        "0.##",
                        CultureInfo.InvariantCulture);

                inbll.discount =
                    discount.ToString(
                        "0.##",
                        CultureInfo.InvariantCulture);

                inbll.due_amount =
                    due.ToString(
                        "0.##",
                        CultureInfo.InvariantCulture);

                inbll.change_amount =
                    change.ToString(
                        "0.##",
                        CultureInfo.InvariantCulture);

                inbll.sales_date = DateTime.Now;

                inbll.added_by = usr.id;

                // ---------------------------------------------
                // INSERT INVOICE
                // ---------------------------------------------

                bool invoiceInserted =
                    indal.insert(inbll);

                if (!invoiceInserted)
                {
                    MessageBox.Show(
                        "Invoice could not be saved.\n\n" +
                        "The invoice record was rejected by the database.",
                        "Transaction Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                // ---------------------------------------------
                // INSERT PRODUCTS
                // ---------------------------------------------

                bool allProductsSaved = true;

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    DataRow row = table.Rows[i];

                    int productId =
                        Convert.ToInt32(row["ProductId"]);

                    string productName =
                        Convert.ToString(row["Item Name"]);

                    string productCode =
                        Convert.ToString(row["Code"]);

                    string productType =
                        Convert.ToString(row["Type"]);

                    decimal salePrice =
                        Convert.ToDecimal(row["Price"]);

                    decimal saleQuantity =
                        Convert.ToDecimal(row["Quantity"]);

                    decimal itemTotal =
                        Convert.ToDecimal(row["Total"]);

                    // -----------------------------------------
                    // GET PRODUCT
                    // -----------------------------------------

                    Manage_Productsbll product =
                        dal.GetProductIDFromName(
                            productName,
                            productType,
                            productCode
                        );

                    if (product == null)
                    {
                        MessageBox.Show(
                            "Product not found in database:\n" +
                            productName,
                            "Product Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        allProductsSaved = false;
                        break;
                    }

                    // Use ProductId returned by database
                    productId = product.id;

                    // -----------------------------------------
                    // GET PURCHASE PRICE
                    // -----------------------------------------

                    Manage_Productsbll purchaseProduct =
                        dal.GetProductpriceFromId(
                            productId.ToString()
                        );

                    if (purchaseProduct == null)
                    {
                        MessageBox.Show(
                            "Purchase price could not be found for:\n" +
                            productName,
                            "Product Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        allProductsSaved = false;
                        break;
                    }

                    decimal purchasePrice;

                    if (!TryParseDecimal(
                        purchaseProduct.purchase_price,
                        out purchasePrice))
                    {
                        MessageBox.Show(
                            "Invalid purchase price for:\n" +
                            productName,
                            "Product Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        allProductsSaved = false;
                        break;
                    }

                    // -----------------------------------------
                    // CALCULATE DISCOUNTED ITEM TOTAL
                    // -----------------------------------------

                    decimal discountAmount =
                        itemTotal * discount / 100;

                    decimal finalItemTotal =
                        itemTotal - discountAmount;

                    // -----------------------------------------
                    // PURCHASE TOTAL
                    // -----------------------------------------

                    decimal purchaseTotal =
                        saleQuantity * purchasePrice;

                    // -----------------------------------------
                    // SET PRODUCT DETAILS
                    // -----------------------------------------

                    prbll.product_id = productId;

                    prbll.product_name =
                        productName;

                    // IMPORTANT:
                    // Do NOT convert "blk" to int.
                    //
                    // If your product_details table's code
                    // column is INT, use productId.
                    //
                    // If code column is VARCHAR/NVARCHAR,
                    // change product_detailsbll.code to string
                    // and use productCode.
                    //
                    // Based on your existing BLL, code appears
                    // to be INT, so Product ID is used here.
                    prbll.code = productId;

                    prbll.type =
                        productType;

                    prbll.price =
                        salePrice.ToString(
                            "0.##",
                            CultureInfo.InvariantCulture);

                    prbll.quantity =
                        saleQuantity.ToString(
                            "0.##",
                            CultureInfo.InvariantCulture);

                    prbll.discount =
                        discount.ToString(
                            "0.##",
                            CultureInfo.InvariantCulture);

                    prbll.total =
                        finalItemTotal.ToString(
                            "0.##",
                            CultureInfo.InvariantCulture);

                    prbll.purchase_price =
                        purchaseTotal.ToString(
                            "0.##",
                            CultureInfo.InvariantCulture);

                    prbll.inv_no =
                        invoiceNumber.ToString();

                    prbll.added_by =
                        usr.id;

                    prbll.added_date =
                        DateTime.Now;

                    // -----------------------------------------
                    // CHECK STOCK BEFORE DECREASING
                    // -----------------------------------------

                    bool stockUpdated =
                        dal.DecreaseProduct(
                            productId,
                            saleQuantity
                        );

                    if (!stockUpdated)
                    {
                        MessageBox.Show(
                            "Stock could not be decreased for:\n" +
                            productName +
                            "\n\nThe transaction was stopped.",
                            "Stock Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        allProductsSaved = false;
                        break;
                    }

                    // -----------------------------------------
                    // INSERT SALE DETAIL
                    // -----------------------------------------

                    bool detailInserted =
                        prdal.insert(prbll);

                    if (!detailInserted)
                    {
                        MessageBox.Show(
                            "Sale detail could not be saved for:\n" +
                            productName +
                            "\n\nThe transaction was stopped.",
                            "Database Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );

                        allProductsSaved = false;
                        break;
                    }
                }

                // ---------------------------------------------
                // FINAL RESULT
                // ---------------------------------------------

                if (!allProductsSaved)
                {
                    MessageBox.Show(
                        "The sale could not be completed.\n\n" +
                        "Please check the database and stock.",
                        "Transaction Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                    return;
                }

                // ---------------------------------------------
                // PRINT
                // ---------------------------------------------

                if (printInvoice)
                {
                    PrintInvoice();
                }

                MessageBox.Show(
                    "Transaction Completed Successfully.",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                // ---------------------------------------------
                // CLEAR AFTER SUCCESS
                // ---------------------------------------------

                ClearAfterSuccessfulSale();

                // Refresh stock
                LoadStock();

                // New invoice number
                ToGetInvoiceID();
            }
            catch (FormatException ex)
            {
                MessageBox.Show(
                    "Number format error.\n\n" +
                    ex.Message +
                    "\n\nPlease check price, quantity, discount and payment values.",
                    "Format Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    "Database error.\n\n" +
                    ex.Message,
                    "Database Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Transaction failed.\n\n" +
                    ex.Message,
                    "Transaction Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // =========================================================
        // PRINT INVOICE
        // =========================================================

        private void PrintInvoice()
        {
            try
            {
                DGVPrinter printer = new DGVPrinter();

                printer.Title =
                    "\r\nTime Paint & Hardware Store";

                printer.SubTitle =
                    "Khokhar Plaza, Near Allied Bank, Main Bahria Enclave Road\r\n" +
                    "Phone: 031659007044\r\n\r\n" +
                    "Invoice #: " +
                    lblinvoiceno.Text +
                    "\r\n";

                printer.SubTitleFormatFlags =
                    StringFormatFlags.LineLimit |
                    StringFormatFlags.NoClip;

                printer.PageNumbers = true;
                printer.PageNumberInHeader = false;
                printer.PorportionalColumns = true;

                printer.HeaderCellAlignment =
                    StringAlignment.Near;

                printer.Footer =
                    "Grand Total: " +
                    txtsubtotal.Text +
                    "\r\n\r\n" +
                    "Total Paid: " +
                    txtpaidamount.Text +
                    "\r\n\r\n" +
                    "Thank You for doing business with us.";

                printer.FooterSpacing = 15;

                printer.PrintDataGridView(
                    dgvinvoice
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Sale was saved, but printing failed.\n\n" +
                    ex.Message,
                    "Printing Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        // =========================================================
        // CLEAR AFTER SUCCESSFUL SALE
        // =========================================================

        private void ClearAfterSuccessfulSale()
        {
            table.Clear();

            dgvinvoice.DataSource = null;
            dgvinvoice.DataSource = table;

            ConfigureInvoiceGrid();

            SetDefaultValues();

            ClearErrors();
        }

        // =========================================================
        // VALIDATE COMPLETE SALE
        // =========================================================

        private bool Validation1()
        {
            ClearErrors();

            if (table == null ||
                table.Rows.Count == 0)
            {
                MessageBox.Show(
                    "Please add at least one product to the invoice.",
                    "Empty Invoice",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            decimal subtotal;

            if (!TryParseDecimal(
                txtsubtotal.Text,
                out subtotal) ||
                subtotal <= 0)
            {
                MessageBox.Show(
                    "Total payable must be greater than zero.",
                    "Invalid Total",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            if (string.IsNullOrWhiteSpace(
                cmbcustomer.Text))
            {
                errorProvider1.SetError(
                    cmbcustomer,
                    "Customer name required."
                );

                MessageBox.Show(
                    "Please select a customer.",
                    "Customer Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            decimal paid;

            if (!TryParseDecimal(
                txtpaidamount.Text,
                out paid))
            {
                MessageBox.Show(
                    "Please enter a valid paid amount.",
                    "Invalid Payment",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            if (paid < 0)
            {
                MessageBox.Show(
                    "Paid amount cannot be negative.",
                    "Invalid Payment",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false;
            }

            return true;
        }

        // =========================================================
        // DECIMAL PARSER
        // =========================================================

        private bool TryParseDecimal(
            string value,
            out decimal result)
        {
            result = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            return decimal.TryParse(
                value.Trim(),
                NumberStyles.Any,
                CultureInfo.CurrentCulture,
                out result
            );
        }

        // =========================================================
        // CLEAR ERRORS
        // =========================================================

        private void ClearErrors()
        {
            if (errorProvider1 != null)
            {
                errorProvider1.Clear();
            }
        }

        // =========================================================
        // GET NEXT INVOICE NUMBER
        // =========================================================

        public bool ToGetInvoiceID()
        {
            bool isSuccess = false;

            connclass c = new connclass();

            using (SqlConnection conn =
                new SqlConnection(c.connection))
            {
                try
                {
                    string query =
                        "SELECT ISNULL(MAX(inv_no), 0) FROM tbl_invoice";

                    using (SqlCommand cmd =
                        new SqlCommand(query, conn))
                    {
                        conn.Open();

                        object result =
                            cmd.ExecuteScalar();

                        int lastInvoice = 0;

                        if (result != null &&
                            result != DBNull.Value)
                        {
                            int.TryParse(
                                result.ToString(),
                                out lastInvoice
                            );
                        }

                        lblinvoiceno.Text =
                            (lastInvoice + 1).ToString();

                        isSuccess = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Unable to get invoice number.\n\n" +
                        ex.Message,
                        "Invoice Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }

            return isSuccess;
        }

        // =========================================================
        // FORM CLOSED
        // =========================================================

        private void Sale_Register_FormClosed(
            object sender,
            FormClosedEventArgs e)
        {
            _instance = null;
        }

        // =========================================================
        // EMPTY EVENTS - KEEP IF WIRED IN DESIGNER
        // =========================================================

        private void dgvinvoice_CellContentClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
        }

        private void dgvinvoice_CellValueNeeded(
            object sender,
            DataGridViewCellValueEventArgs e)
        {
        }

        private void txttotal_TextChanged(
            object sender,
            EventArgs e)
        {
        }

        private void dgvinvoice_RowHeaderMouseClick(
            object sender,
            DataGridViewCellMouseEventArgs e)
        {
        }

        private void txtprice_TextChanged(
            object sender,
            EventArgs e)
        {
        }

        private void txtdueamount_TextChanged(
            object sender,
            EventArgs e)
        {
        }

        private void cmbcustomer_SelectedIndexChanged(
            object sender,
            EventArgs e)
        {
        }

        private void txtchangeamount_TextChanged(
            object sender,
            EventArgs e)
        {
        }

        private void printDocument1_PrintPage(
            object sender,
            System.Drawing.Printing.PrintPageEventArgs e)
        {
        }
    }
}