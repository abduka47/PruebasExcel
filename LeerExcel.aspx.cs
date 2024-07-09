using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PruebasExcel
{
    public partial class LeerExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (fuExcel.HasFile)
            {
                string userFolderPath = @"C:\\Users\\abdu.rieder\\Downloads";
                string filePath = Path.Combine(userFolderPath, fuExcel.FileName);
                try
                {
                    fuExcel.SaveAs(filePath);
                    int sheetNumber;
                    if (int.TryParse(txtSheetNumber.Text, out sheetNumber) && sheetNumber > 0)
                    {
                        string htmlTable = ReadExcelFile(filePath, sheetNumber);
                        ltExcelTable.Text = htmlTable;
                    }
                    else
                    {
                        ltExcelTable.Text = "Por favor, ingrese un número de hoja válido.";
                    }
                }

                catch (Exception ex)
                {
                    Response.Write($"<script>alert('Error: {ex.Message}');</script>");
                }
            }
        }
        private string ReadExcelFile(string filePath, int sheetNumber)
        {
            StringBuilder htmlTable = new StringBuilder();
            htmlTable.Append("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>");

            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                if (sheetNumber > workbook.Worksheets.Count)
                {
                    return "El número de hoja ingresado excede el número de hojas en el archivo.";
                }
                IXLWorksheet worksheet = workbook.Worksheet(sheetNumber); // Lee la hoja especificada
                foreach (IXLRow row in worksheet.RowsUsed()) // Solo filas usadas
                {
                    htmlTable.Append("<tr>");
                    foreach (IXLCell cell in row.Cells()) // Todas las celdas de la fila
                    {
                        htmlTable.Append("<td style='border: 1px solid black; padding: 5px;'>");
                        try
                        {
                            if (!cell.IsEmpty())
                            {
                                string cellValue = cell.Value.ToString();
                                if (double.TryParse(cellValue, out double numericValue))
                                {
                                    //Formatear valores numéricos con puntos de miles
                                    string formattedValue = numericValue.ToString("N0", CultureInfo.GetCultureInfo("es-CO"));
                                    htmlTable.Append(formattedValue);
                                }
                                else
                                {
                                    htmlTable.Append(cellValue);
                                }
                            }
                        }
                        catch

                        {
                            //htmlTable.Append($"[Error: {ex.Message}]");
                            htmlTable.Append("");
                        }
                        htmlTable.Append("</td>");
                    }
                    htmlTable.Append("</tr>");
                }
            }
            htmlTable.Append("</table>");
            return htmlTable.ToString();
        }
    }
}
