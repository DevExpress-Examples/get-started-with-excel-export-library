#Region "#Namespaces"
Imports DevExpress.Export.Xl
Imports System.IO
' ...
#End Region ' #Namespaces

#Region "#Code"
Namespace XLExportExamples
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            ' Create an exporter instance. 
            Dim exporter As IXlExporter = XlExport.CreateExporter(XlDocumentFormat.Xlsx)

            ' Create the FileStream object with the specified file path. 
            Using stream As New FileStream("Document.xlsx", FileMode.Create, FileAccess.ReadWrite)

                ' Create a new document and begin to write it to the specified stream. 
                Using document As IXlDocument = exporter.CreateDocument(stream)

                    ' Add a new worksheet to the document. 
                    Using sheet As IXlSheet = document.CreateSheet()

                        ' Specify the worksheet name.
                        sheet.Name = "Sales report"

                        ' Create the first column and set its width. 
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                        End Using

                        ' Create the second column and set its width.
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 250
                        End Using

                        ' Create the third column and set the specific number format for its cells.
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using

                        ' Specify cell font attributes.
                        Dim cellFormatting As New XlCellFormatting()
                        cellFormatting.Font = New XlFont()
                        cellFormatting.Font.Name = "Century Gothic"
                        cellFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

                        ' Specify formatting settings for the header row.
                        Dim headerRowFormatting As New XlCellFormatting()
                        headerRowFormatting.CopyFrom(cellFormatting)
                        headerRowFormatting.Font.Bold = True
                        headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                        headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))

                        ' Create the header row.
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = "Region"
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = "Product"
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = "Sales"
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        End Using

                        ' Generate data for the sales report.
                        Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" }
                        Dim amount() As Integer = { 6750, 4500, 3550, 4250, 5500, 6250, 5325, 4235 }
                        For i As Integer = 0 To 7
                            Using row As IXlRow = sheet.CreateRow()
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = If(i < 4, "East", "West")
                                    cell.ApplyFormatting(cellFormatting)
                                End Using
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = products(i Mod 4)
                                    cell.ApplyFormatting(cellFormatting)
                                End Using
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = amount(i)
                                    cell.ApplyFormatting(cellFormatting)
                                End Using
                            End Using
                        Next i

                        ' Enable AutoFilter for the created cell range.
                        sheet.AutoFilterRange = sheet.DataRange

                        ' Specify formatting settings for the total row.
                        Dim totalRowFormatting As New XlCellFormatting()
                        totalRowFormatting.CopyFrom(cellFormatting)
                        totalRowFormatting.Font.Bold = True
                        totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent5, 0.6))

                        ' Create the total row.
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.ApplyFormatting(totalRowFormatting)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = "Total amount"
                                cell.ApplyFormatting(totalRowFormatting)
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom))
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                ' Add values in the cell range C2 through C9 using the SUBTOTAL function. 
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(2, 1, 2, 8), XlSummary.Sum, True))
                                cell.ApplyFormatting(totalRowFormatting)
                            End Using
                        End Using
                    End Using
                End Using
            End Using
            ' Open the XLSX document using the default application.
            System.Diagnostics.Process.Start("Document.xlsx")
        End Sub
    End Class
End Namespace
#End Region ' #Code