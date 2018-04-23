#region #Namespaces
using DevExpress.Export.Xl;
using System.IO;
// ...
#endregion #Namespaces

#region #Code
namespace XLExportExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an exporter instance. 
            IXlExporter exporter = XlExport.CreateExporter(XlDocumentFormat.Xlsx);

            // Create the FileStream object with the specified file path. 
            using (FileStream stream = new FileStream("Document.xlsx", FileMode.Create, FileAccess.ReadWrite)) {
                
                // Create a new document and begin to write it to the specified stream. 
                using (IXlDocument document = exporter.CreateDocument(stream)) {
                    
                    // Add a new worksheet to the document. 
                    using (IXlSheet sheet = document.CreateSheet()) {  
       
                        // Specify the worksheet name.
                        sheet.Name = "Sales report";

                        // Create the first column and set its width. 
                        using (IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }

                        // Create the second column and set its width.
                        using (IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 250;
                        }

                        // Create the third column and set the specific number format for its cells.
                        using (IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }

                        // Specify cell font attributes.
                        XlCellFormatting cellFormatting = new XlCellFormatting();
                        cellFormatting.Font = new XlFont();
                        cellFormatting.Font.Name = "Century Gothic";
                        cellFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                        // Specify formatting settings for the header row.
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.CopyFrom(cellFormatting);
                        headerRowFormatting.Font.Bold = true;
                        headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                        headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));

                        // Create the header row.
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = "Region";
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = "Product";
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = "Sales";
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }

                        // Generate data for the sales report.
                        string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                        int[] amount = new int[] { 6750, 4500, 3550, 4250, 5500, 6250, 5325, 4235 };
                        for (int i = 0; i < 8; i++)
                        {
                            using (IXlRow row = sheet.CreateRow()) {
                                using (IXlCell cell = row.CreateCell()) {
                                    cell.Value = (i < 4) ? "East" : "West";
                                    cell.ApplyFormatting(cellFormatting);
                                }
                                using (IXlCell cell = row.CreateCell()) {
                                    cell.Value = products[i % 4];
                                    cell.ApplyFormatting(cellFormatting);
                                }
                                using (IXlCell cell = row.CreateCell()) {
                                    cell.Value = amount[i];
                                    cell.ApplyFormatting(cellFormatting);
                                }
                            }
                        }

                        // Enable AutoFilter for the created cell range.
                        sheet.AutoFilterRange = sheet.DataRange;

                        // Specify formatting settings for the total row.
                        XlCellFormatting totalRowFormatting = new XlCellFormatting();
                        totalRowFormatting.CopyFrom(cellFormatting);
                        totalRowFormatting.Font.Bold = true;
                        totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent5, 0.6));

                        // Create the total row.
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) { 
                                cell.ApplyFormatting(totalRowFormatting); 
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = "Total amount";
                                cell.ApplyFormatting(totalRowFormatting);
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Add values in the cell range C2 through C9 using the SUBTOTAL function. 
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(2, 1, 2, 8), XlSummary.Sum, true));
                                cell.ApplyFormatting(totalRowFormatting);
                            }
                        }
                    }
                }
            }
            // Open the XLSX document using the default application.
            System.Diagnostics.Process.Start("Document.xlsx");
        }
    }
}
#endregion #Code