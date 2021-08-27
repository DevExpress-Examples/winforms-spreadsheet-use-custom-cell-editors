<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613964/19.2.2%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T419595)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/DevAVInvoicing/Form1.cs) (VB: [Form1.vb](./VB/DevAVInvoicing/Form1.vb))
<!-- default file list end -->
# How to use custom cell editors to create a data entry form


This example demonstrates how to use <a href="https://documentation.devexpress.com/#WindowsForms/CustomDocument18170">custom cell editors</a> to create a data entry form that allows end-users to quickly generate invoices. The required data entry fields are marked with an asterisk. To add a new record to the invoice or delete the existing one, an end-user should switch to the <strong>Invoice</strong> tab and click the <strong>Add</strong> or <strong>Remove</strong> button, respectively. All other content is protected to prevent inappropriate modifications.<br><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-use-custom-cell-editors-to-create-a-data-entry-form-t419595/16.1.5+/media/69f65ea6-6b87-11e6-80bf-00155d62480c.png"><br>Data for the invoice is provided based on the document template (<em>DevAVInvoicing.xltx</em>), which includes the following worksheets:<br>• Invoice – contains a sales entry form;<br>• Customers (hidden) – contains customer info;<br>• Employees (hidden) – contains a list of employees;<br>• Products (hidden) – contains product data;<br>• Stores (hidden) – contains information about stores owned by customers.<br><br>To retrieve the required data from worksheets, the Spreadsheet uses the <a href="https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1">VLOOKUP</a> and <a href="https://support.office.com/en-us/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e">DGET</a> functions. For example, when an end-user selects a customer's name in the Billing Address section, <strong>VLOOKUP</strong> is used to find and display the customer's billing address. In the same way, the <strong>DGET</strong> function is used to automatically display the shipping address based on the customer's name and the selected city of the store to which the order should be delivered.

<br/>


