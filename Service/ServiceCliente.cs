using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using RoboBarbearia.Model;
using RoboBarbearia.Properties;
using RoboBarbearia.Utils;

namespace RoboBarbearia.Service
{
    public static class ServiceCliente
    {
        public static List<Cliente> BuscarClientes()
        {
            List<Cliente> clienteLista = null;
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Usuarios.xlsx"));
                var workBook = package.Workbook;
                clienteLista = new List<Cliente>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Planilha1"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 4].Value.ToString()))
                            clienteLista.Add(new Cliente(
                                worksheet.Cells[row, 1].Value?.ToString().Trim(),
                                worksheet.Cells[row, 5].Value?.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 6].Value?.ToString().Trim(),
                                worksheet.Cells[row, 7].Value?.ToString().Trim(),
                                worksheet.Cells[row, 8].Value?.ToString().Trim(),
                                worksheet.Cells[row, 9].Value?.ToString().Trim(),
                                worksheet.Cells[row, 10].Value?.ToString(),
                                worksheet.Cells[row, 12].Value?.ToString(),
                                worksheet.Cells[row, 13].Value?.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 14].Value?.ToString(),
                                worksheet.Cells[row, 15].Value?.ToString()
                            ));
                }

                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarClientes", ex);
            }
            return new List<Cliente>(clienteLista ?? throw new InvalidOperationException("Rotina BuscarClientes, retornou null!"));
        }
        
        public static void AtualizarCliente(string pNomeCliente, bool xEhErro, bool xEhRelatorio, bool xEhFinanceiro)
        {
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Usuarios.xlsx"));
                var workBook = package.Workbook;

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Planilha1"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                    {
                        if (xEhRelatorio)
                        {
                            if (worksheet.Cells[row, 1].Value.ToString() == pNomeCliente)
                                worksheet.Cells[row, 12].Value = xEhErro ? "ERRO" : DateTime.Now.ToString("ddMMyy");
                        } 
                        else if (xEhFinanceiro) 
                        {
                            if (worksheet.Cells[row, 1].Value.ToString() == pNomeCliente)
                                worksheet.Cells[row, 15].Value = xEhErro ? "ERRO" : DateTime.Now.ToString("ddMMyy");
                        }                    
                    }
                }

                package.Save();
                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarClientes", ex);
            }
        }
    }
}