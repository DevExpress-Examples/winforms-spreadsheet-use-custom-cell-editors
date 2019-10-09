Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraBars
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraSpreadsheet

Namespace DevAVInvoicing
	Partial Public Class Form1
		Inherits RibbonForm

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
			LoadInvoiceTemplate()
			ribbonControl1.Minimized = True
		End Sub

		Private Sub LoadInvoiceTemplate()
			Dim fileName As String = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "DevAVInvoicing.xltx")
			spreadsheetControl1.LoadDocument(fileName)
			BindCustomEditors()
		End Sub

		Private Sub BindCustomEditors()
			Dim workbook As IWorkbook = spreadsheetControl1.Document
			Dim invoice As Worksheet = workbook.Worksheets("Invoice")

			' Use a combo box editor as the in-place editor for cells containing a customer's name in the Billing Address section.
			' The editor's items are obtained from the "Customers" worksheet.
			Dim customers As Worksheet = workbook.Worksheets("Customers")
			invoice.CustomCellInplaceEditors.Add(invoice("B10:C10"), CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(customers("A2:A21")))

			' Use a combo box editor as the in-place editor for cells containing the store location in the Shipping Address section.
			' The editor's items are obtained from the "Stores" worksheet.
			Dim stores As Worksheet = workbook.Worksheets("Stores")
			invoice.CustomCellInplaceEditors.Add(invoice("G12:I12"), CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(stores("D5:D204")), True)

			' Use a combo box editor as the in-place editor for cells containing a sales representative's name in the "Sales Rep." column.
			' The editor's items are obtained from the "Employees" worksheet.
			Dim employees As Worksheet = workbook.Worksheets("Employees")
			invoice.CustomCellInplaceEditors.Add(invoice("B18:C18"), CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(employees("I2:I52")), True)

			' Use a date editor as the in-place editor for cells containing the shipping date in the "Ship date" column.
			invoice.CustomCellInplaceEditors.Add(invoice("F18:G18"), CustomCellInplaceEditorType.DateEdit)

			' Use a combo box editor as the in-place editor for cells allowing a user to select the preferred shipping method in the "Ship via" column.
			Dim shipVia() As CellValue = { "Air", "Ground", "Ship" }
			invoice.CustomCellInplaceEditors.Add(invoice("H18:I18"), CustomCellInplaceEditorType.ComboBox, ValueObject.CreateListSource(shipVia))

			' Use the custom control (SpinEdit) as the in-place editor for cells containing the FOB value.
			' To provide the editor, handle the CustomCellEdit event. 
			invoice.CustomCellInplaceEditors.Add(invoice("J18:K18"), CustomCellInplaceEditorType.Custom, "FOBSpinEdit")

			' Use the custom control (SpinEdit) as the in-place editor for cells containing the invoice payment terms.
			' To provide the editor, handle the CustomCellEdit event. 
			invoice.CustomCellInplaceEditors.Add(invoice("L18:M18"), CustomCellInplaceEditorType.Custom, "TermsSpinEdit")

			' Use the custom control (SpinEdit) as the in-place editor for cells containing the product quantity in the "Quantity" column.
			' To provide the editor, handle the CustomCellEdit event. 
			invoice.CustomCellInplaceEditors.Add(invoice("B22:B25"), CustomCellInplaceEditorType.Custom, "QtySpinEdit")

			' Use a combo box editor as the in-place editor for cells containing product names in the "Description" column.
			' The editor's items are obtained from the "Products" worksheet.
			Dim products As Worksheet = workbook.Worksheets("Products")
			invoice.CustomCellInplaceEditors.Add(invoice("C22:F25"), CustomCellInplaceEditorType.ComboBox, ValueObject.FromRange(products("B2:B20")))

			' Use the custom control (SpinEdit) as the in-place editor for cells containing the discount value in the "Discount" column.
			' To provide the editor, handle the CustomCellEdit event. 
			invoice.CustomCellInplaceEditors.Add(invoice("I22:J25"), CustomCellInplaceEditorType.Custom, "DiscountSpinEdit")

			' Use the custom control (SpinEdit) as the in-place editor for cells containing the shipping costs.
			' To provide the editor, handle the CustomCellEdit event. 
			invoice.CustomCellInplaceEditors.Add(invoice("K27:M27"), CustomCellInplaceEditorType.Custom, "ShippingSpinEdit")
		End Sub

		Private Sub spreadsheetControl1_CustomCellEdit(ByVal sender As Object, ByVal e As SpreadsheetCustomCellEditEventArgs) Handles spreadsheetControl1.CustomCellEdit
			If Not e.ValueObject.IsText Then
				Return
			End If
			If e.ValueObject.TextValue = "FOBSpinEdit" Then
				e.RepositoryItem = CreateSpinEdit(0, 500, 5)
			ElseIf e.ValueObject.TextValue = "TermsSpinEdit" Then
				e.RepositoryItem = CreateSpinEdit(5, 30, 1)
			ElseIf e.ValueObject.TextValue = "QtySpinEdit" Then
				e.RepositoryItem = CreateSpinEdit(1, 100, 1)
			ElseIf e.ValueObject.TextValue = "DiscountSpinEdit" Then
				e.RepositoryItem = CreateSpinEdit(0, 1000, 10)
			ElseIf e.ValueObject.TextValue = "ShippingSpinEdit" Then
				e.RepositoryItem = CreateSpinEdit(10, 1000, 5)
			End If
		End Sub

		Private Function CreateSpinEdit(ByVal minValue As Integer, ByVal maxValue As Integer, ByVal increment As Integer) As RepositoryItemSpinEdit
			Dim editor As New RepositoryItemSpinEdit()
			editor.AutoHeight = False
			editor.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
			editor.MinValue = minValue
			editor.MaxValue = maxValue
			editor.Increment = increment
			editor.IsFloatValue = False
			Return editor
		End Function

		Private Sub spreadsheetControl1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles spreadsheetControl1.SelectionChanged
			EnableControls()
			ActivateEditor()
		End Sub

		Private Sub EnableControls()
			Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
			If sheet.Name = "Invoice" Then
				Dim invoiceItems As DefinedName = sheet.DefinedNames.GetDefinedName("InvoiceItems")
				btnRemoveRecord.Enabled = invoiceItems IsNot Nothing AndAlso invoiceItems.Range.RowCount > 1 AndAlso invoiceItems.Range.IsIntersecting(sheet.SelectedCell)
			Else
				btnRemoveRecord.Enabled = False
			End If
		End Sub

		Private Sub ActivateEditor()
			Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
			If sheet.Name = "Invoice" Then
				Dim editors As IList(Of CustomCellInplaceEditor) = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection)
				If editors.Count = 1 Then
					spreadsheetControl1.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit)
				End If
			End If
		End Sub

		Private Sub spreadsheetControl1_CellValueChanged(ByVal sender As Object, ByVal e As SpreadsheetCellEventArgs) Handles spreadsheetControl1.CellValueChanged
			If e.Action = CellValueChangedAction.UndoRedo OrElse e.OldValue = e.Cell.Value OrElse e.Cell.GetReferenceA1(ReferenceElement.IncludeSheetName) <> "Invoice!B10" Then
				Return
			End If
			Dim invoice As Worksheet = e.Worksheet
			Dim customerStores As Worksheet = spreadsheetControl1.Document.Worksheets("Stores")
			' Apply a filter to the "CustomerId" column of the "StoresTable" table on the "Stores" worksheet 
			' to display stores owned by the customer with the specified ID.
			Dim customerId As String = invoice("B11").Value.TextValue
			Dim storesTable As Table = customerStores.Tables(0)
			storesTable.AutoFilter.Clear()
			storesTable.AutoFilter.Columns(1).ApplyFilterCriteria(customerId)
			' Select the default store and assign it to the cell G12 on the "Invoice" worksheet.
			Dim range As CellRange = storesTable.DataRange
			For rowOffset As Integer = 0 To range.RowCount - 1
				If range(rowOffset, 1).Value.TextValue = customerId Then
					invoice("G12").Value = range(rowOffset, 3).Value.TextValue
					Return
				End If
			Next rowOffset
			invoice("G12").Value = CellValue.Empty
		End Sub

		Private Sub btnRemoveRecord_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles btnRemoveRecord.ItemClick
			Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
			If sheet.Name = "Invoice" Then
				If spreadsheetControl1.IsCellEditorActive Then
					spreadsheetControl1.CloseCellEditor(CellEditorEnterValueMode.Cancel)
				End If
				sheet.Rows.Remove(sheet.SelectedCell.TopRowIndex, 1)
				EnableControls()
				ActivateEditor()
			End If
		End Sub

		Private Sub btnAddRecord_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles btnAddRecord.ItemClick
			Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
			If sheet.Name = "Invoice" Then
				If spreadsheetControl1.IsCellEditorActive Then
					spreadsheetControl1.CloseCellEditor(CellEditorEnterValueMode.Cancel)
				End If
				AddRecord(sheet)
				EnableControls()
				ActivateEditor()
			End If
		End Sub

		' Add a new record to the invoice.
		Private Sub AddRecord(ByVal sheet As Worksheet)
			spreadsheetControl1.BeginUpdate()
			Try
				Dim invoiceItems As DefinedName = sheet.DefinedNames.GetDefinedName("InvoiceItems")
				Dim rowIndex As Integer = invoiceItems.Range.BottomRowIndex
				sheet.Rows.Insert(rowIndex)
				sheet.Rows(rowIndex).Height = sheet.Rows(rowIndex + 1).Height
				Dim range As CellRange = invoiceItems.Range
				Dim itemRange As CellRange = sheet.Range.FromLTRB(range.LeftColumnIndex, range.BottomRowIndex, range.RightColumnIndex, range.BottomRowIndex)
				MoveUpLastRecord(itemRange)
				InitializeRecord(itemRange)
				spreadsheetControl1.SelectedCell = itemRange(1)
			Finally
				spreadsheetControl1.EndUpdate()
			End Try
		End Sub

		' Move the last record one row up.
		Private Sub MoveUpLastRecord(ByVal itemRange As CellRange)
			Dim range As CellRange = itemRange.Offset(-1, 0)
			range.CopyFrom(itemRange, PasteSpecial.All, True)
		End Sub

		' Specify the default values for a new record. 
		Private Sub InitializeRecord(ByVal itemRange As CellRange)
			itemRange(0).Value = 1 ' Quantity
			itemRange(1).Value = CellValue.Empty ' Product Description
			itemRange(7).Value = 0 ' Discount
		End Sub

		' Suppress the protection warning dialog.
		Private Sub spreadsheetControl1_ProtectionWarning(ByVal sender As Object, ByVal e As HandledEventArgs) Handles spreadsheetControl1.ProtectionWarning
			e.Handled = True
		End Sub

		' Load the invoice template when a new document is created.
		Private Sub spreadsheetControl1_EmptyDocumentCreated(ByVal sender As Object, ByVal e As EventArgs) Handles spreadsheetControl1.EmptyDocumentCreated
			LoadInvoiceTemplate()
		End Sub

		' Handle the CustomDrawCell event to mark data entry fields with an asterisk.
		Private Sub spreadsheetControl1_CustomDrawCell(ByVal sender As Object, ByVal e As CustomDrawCellEventArgs) Handles spreadsheetControl1.CustomDrawCell
			Dim cellReference As String = e.Cell.GetReferenceA1()
			If e.Cell.Worksheet.Name <> "Invoice" OrElse (cellReference <> "A5" AndAlso cellReference <> "A10" AndAlso cellReference <> "F12") Then
				Return
			End If
			e.Handled = True
			e.DrawDefault()
			Using font As New Font(e.Font.Name, 14F, FontStyle.Bold)
				Dim text As String = "*"
				Dim size As SizeF = e.Graphics.MeasureString(text, font, Int32.MaxValue, StringFormat.GenericDefault)
				Dim textBounds As New RectangleF(e.Bounds.Right - size.Width - 2, e.Bounds.Bottom - size.Height * 0.7F, size.Width + 2, size.Height)
				e.Graphics.DrawString(text, font, e.Cache.GetSolidBrush(Color.Red), textBounds, StringFormat.GenericDefault)
			End Using
		End Sub
	End Class
End Namespace
