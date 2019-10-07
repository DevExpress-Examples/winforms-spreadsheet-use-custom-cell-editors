using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraSpreadsheet;

namespace DevAVInvoicing
{
    public partial class Form1 : RibbonForm {
        public Form1() {
            InitializeComponent();
        }

        void Form1_Load(object sender, EventArgs e) {
            LoadInvoiceTemplate();
            ribbonControl1.Minimized = true;
        }

        void LoadInvoiceTemplate() {
            string fileName = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "DevAVInvoicing.xltx");
            spreadsheetControl1.LoadDocument(fileName);
            BindCustomEditors();
        }

        void BindCustomEditors() {
            IWorkbook workbook = spreadsheetControl1.Document;
            Worksheet invoice = workbook.Worksheets["Invoice"];

            // Use a combo box editor as the in-place editor for cells containing a customer's name in the Billing Address section.
            // The editor's items are obtained from the "Customers" worksheet.
            Worksheet customers = workbook.Worksheets["Customers"];
            invoice.CustomCellInplaceEditors.Add(invoice["B10:C10"], CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(customers["A2:A21"]));

            // Use a combo box editor as the in-place editor for cells containing the store location in the Shipping Address section.
            // The editor's items are obtained from the "Stores" worksheet.
            Worksheet stores = workbook.Worksheets["Stores"];
            invoice.CustomCellInplaceEditors.Add(invoice["G12:I12"], CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(stores["D5:D204"]), true);

            // Use a combo box editor as the in-place editor for cells containing a sales representative's name in the "Sales Rep." column.
            // The editor's items are obtained from the "Employees" worksheet.
            Worksheet employees = workbook.Worksheets["Employees"];
            invoice.CustomCellInplaceEditors.Add(invoice["B18:C18"], CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(employees["I2:I52"]), true);

            // Use a date editor as the in-place editor for cells containing the shipping date in the "Ship date" column.
            invoice.CustomCellInplaceEditors.Add(invoice["F18:G18"], CustomCellInplaceEditorType.DateEdit);
            
            // Use a combo box editor as the in-place editor for cells allowing a user to select the preferred shipping method in the "Ship via" column.
            CellValue[] shipVia = { "Air", "Ground", "Ship" };
            invoice.CustomCellInplaceEditors.Add(invoice["H18:I18"], CustomCellInplaceEditorType.ComboBox, ValueObject.CreateListSource(shipVia));

            // Use the custom control (SpinEdit) as the in-place editor for cells containing the FOB value.
            // To provide the editor, handle the CustomCellEdit event. 
            invoice.CustomCellInplaceEditors.Add(invoice["J18:K18"], CustomCellInplaceEditorType.Custom, "FOBSpinEdit");
            
            // Use the custom control (SpinEdit) as the in-place editor for cells containing the invoice payment terms.
            // To provide the editor, handle the CustomCellEdit event. 
            invoice.CustomCellInplaceEditors.Add(invoice["L18:M18"], CustomCellInplaceEditorType.Custom, "TermsSpinEdit");
             
            // Use the custom control (SpinEdit) as the in-place editor for cells containing the product quantity in the "Quantity" column.
            // To provide the editor, handle the CustomCellEdit event. 
            invoice.CustomCellInplaceEditors.Add(invoice["B22:B25"], CustomCellInplaceEditorType.Custom, "QtySpinEdit");

            // Use a combo box editor as the in-place editor for cells containing product names in the "Description" column.
            // The editor's items are obtained from the "Products" worksheet.
            Worksheet products = workbook.Worksheets["Products"];
            invoice.CustomCellInplaceEditors.Add(invoice["C22:F25"], CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(products["B2:B20"]));
            
            // Use the custom control (SpinEdit) as the in-place editor for cells containing the discount value in the "Discount" column.
            // To provide the editor, handle the CustomCellEdit event. 
            invoice.CustomCellInplaceEditors.Add(invoice["I22:J25"], CustomCellInplaceEditorType.Custom, "DiscountSpinEdit");

            // Use the custom control (SpinEdit) as the in-place editor for cells containing the shipping costs.
            // To provide the editor, handle the CustomCellEdit event. 
            invoice.CustomCellInplaceEditors.Add(invoice["K27:M27"], CustomCellInplaceEditorType.Custom, "ShippingSpinEdit");
        }

        void spreadsheetControl1_CustomCellEdit(object sender, SpreadsheetCustomCellEditEventArgs e) {
            if (!e.ValueObject.IsText)
                return;
            if (e.ValueObject.TextValue == "FOBSpinEdit")
                e.RepositoryItem = CreateSpinEdit(0, 500, 5);
            else if (e.ValueObject.TextValue == "TermsSpinEdit")
                e.RepositoryItem = CreateSpinEdit(5, 30, 1);
            else if (e.ValueObject.TextValue == "QtySpinEdit")
                e.RepositoryItem = CreateSpinEdit(1, 100, 1);
            else if (e.ValueObject.TextValue == "DiscountSpinEdit")
                e.RepositoryItem = CreateSpinEdit(0, 1000, 10);
            else if (e.ValueObject.TextValue == "ShippingSpinEdit")
                e.RepositoryItem = CreateSpinEdit(10, 1000, 5);
        }

        RepositoryItemSpinEdit CreateSpinEdit(int minValue, int maxValue, int increment) {
            RepositoryItemSpinEdit editor = new RepositoryItemSpinEdit();
            editor.AutoHeight = false;
            editor.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            editor.MinValue = minValue;
            editor.MaxValue = maxValue;
            editor.Increment = increment;
            editor.IsFloatValue = false;
            return editor;
        }

        void spreadsheetControl1_SelectionChanged(object sender, EventArgs e) { 
            EnableControls();
            ActivateEditor();
        }

        void EnableControls() {
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet.Name == "Invoice") {
                DefinedName invoiceItems = sheet.DefinedNames.GetDefinedName("InvoiceItems");
                btnRemoveRecord.Enabled = invoiceItems != null && invoiceItems.Range.RowCount > 1 && invoiceItems.Range.IsIntersecting(sheet.SelectedCell);
            }
            else
                btnRemoveRecord.Enabled = false;
        }

        void ActivateEditor() {
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet.Name == "Invoice") {
                IList<CustomCellInplaceEditor> editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    spreadsheetControl1.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit);
            }
        }

        void spreadsheetControl1_CellValueChanged(object sender, SpreadsheetCellEventArgs e) {
            if (e.Action == CellValueChangedAction.UndoRedo || e.OldValue == e.Cell.Value || 
                e.Cell.GetReferenceA1(ReferenceElement.IncludeSheetName) != "Invoice!B10")
                return;
            Worksheet invoice = e.Worksheet;
            Worksheet customerStores = spreadsheetControl1.Document.Worksheets["Stores"];
            // Apply a filter to the "CustomerId" column of the "StoresTable" table on the "Stores" worksheet 
            // to display stores owned by the customer with the specified ID.
            string customerId = invoice["B11"].Value.TextValue;
            Table storesTable = customerStores.Tables[0];
            storesTable.AutoFilter.Clear();
            storesTable.AutoFilter.Columns[1].ApplyFilterCriteria(customerId);
            // Select the default store and assign it to the cell G12 on the "Invoice" worksheet.
            CellRange range = storesTable.DataRange;
            for (int rowOffset = 0; rowOffset < range.RowCount; rowOffset++) {
                if (range[rowOffset, 1].Value.TextValue == customerId) {
                    invoice["G12"].Value = range[rowOffset, 3].Value.TextValue;
                    return;
                }
            }
            invoice["G12"].Value = CellValue.Empty;
        }

        void btnRemoveRecord_ItemClick(object sender, ItemClickEventArgs e) {
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet.Name == "Invoice") {
                if (spreadsheetControl1.IsCellEditorActive)
                    spreadsheetControl1.CloseCellEditor(CellEditorEnterValueMode.Cancel);
                sheet.Rows.Remove(sheet.SelectedCell.TopRowIndex, 1);
                EnableControls();
                ActivateEditor();
            }
        }

        void btnAddRecord_ItemClick(object sender, ItemClickEventArgs e) {
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet.Name == "Invoice") {
                if (spreadsheetControl1.IsCellEditorActive)
                    spreadsheetControl1.CloseCellEditor(CellEditorEnterValueMode.Cancel);
                AddRecord(sheet);
                EnableControls();
                ActivateEditor();
            }
        }

        // Add a new record to the invoice.
        void AddRecord(Worksheet sheet) {
            spreadsheetControl1.BeginUpdate();
            try {
                DefinedName invoiceItems = sheet.DefinedNames.GetDefinedName("InvoiceItems");
                int rowIndex = invoiceItems.Range.BottomRowIndex;
                sheet.Rows.Insert(rowIndex);
                sheet.Rows[rowIndex].Height = sheet.Rows[rowIndex + 1].Height;
                CellRange range = invoiceItems.Range;
                CellRange itemRange = sheet.Range.FromLTRB(range.LeftColumnIndex, range.BottomRowIndex, range.RightColumnIndex, range.BottomRowIndex);
                MoveUpLastRecord(itemRange);
                InitializeRecord(itemRange);
                spreadsheetControl1.SelectedCell = itemRange[1];
            }
            finally {
                spreadsheetControl1.EndUpdate();
            }
        }

        // Move the last record one row up.
        void MoveUpLastRecord(CellRange itemRange) {
            CellRange range = itemRange.Offset(-1, 0);
            range.CopyFrom(itemRange, PasteSpecial.All, true);
        }

        // Specify the default values for a new record. 
        void InitializeRecord(CellRange itemRange) {
            itemRange[0].Value = 1; // Quantity
            itemRange[1].Value = CellValue.Empty; // Product Description
            itemRange[7].Value = 0; // Discount
        }

        // Suppress the protection warning dialog.
        void spreadsheetControl1_ProtectionWarning(object sender, HandledEventArgs e) {
            e.Handled = true;
        }

        // Load the invoice template when a new document is created.
        void spreadsheetControl1_EmptyDocumentCreated(object sender, EventArgs e) {
            LoadInvoiceTemplate();
        }

        // Handle the CustomDrawCell event to mark data entry fields with an asterisk.
        void spreadsheetControl1_CustomDrawCell(object sender, CustomDrawCellEventArgs e) {
            string cellReference = e.Cell.GetReferenceA1();
            if (e.Cell.Worksheet.Name != "Invoice" || (cellReference != "A5" && cellReference != "A10" && cellReference != "F12"))
                return;
            e.Handled = true;
            e.DrawDefault();
            using (Font font = new Font(e.Font.Name, 14f, FontStyle.Bold)) {
                string text = "*";
                SizeF size = e.Graphics.MeasureString(text, font, Int32.MaxValue, StringFormat.GenericDefault);
                RectangleF textBounds = new RectangleF(e.Bounds.Right - size.Width - 2, e.Bounds.Bottom - size.Height * 0.7f, size.Width + 2, size.Height);
                e.Graphics.DrawString(text, font, e.Cache.GetSolidBrush(Color.Red), textBounds, StringFormat.GenericDefault);
            }
        }
    }
}
