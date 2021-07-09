using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using RoboBarbearia.Model;
using RoboBarbearia.Properties;
using RoboBarbearia.Utils;

namespace RoboBarbearia.Service
{
    public static class ServiceRelatorio
    {
        public static List<Relatorio> BuscarRelatorios()
        {
            List<Relatorio> relatorioLista = null;
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Relatorios.xlsx"));

                var workBook = package.Workbook;
                relatorioLista = new List<Relatorio>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Relatorios"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 2].Value.ToString()))
                            relatorioLista.Add(new Relatorio(worksheet.Cells[row, 2].Value.ToString().Trim(),
                                worksheet.Cells[row, 3].Value.ToString().Trim(),
                                worksheet.Cells[row, 4].Value.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 5].Value.ToString().Trim()));
                }

                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarRelatorios", ex);
            }
            return new List<Relatorio>(relatorioLista ?? throw new InvalidOperationException("Rotina BuscarRelatorios, retornou null!"));
        }
        
        public static void BaixarRelatorios(Cliente xCliente, IEnumerable<Relatorio> xRelatorioLista)
        {
            bool ErrorDownload = false;

            foreach (var relatorio in xRelatorioLista.Where(relatorio => relatorio.AtivoRelatorio))
            {
                var optionsChr = new ChromeOptions();
                // Roda sem o browser
                //optionsChr.AddArgument("--headless");
                optionsChr.AddArgument("--ignore-certificate-errors");
                optionsChr.AddArgument("--ignore-ssl-errors");

                // Inicializa o Chrome Driver
                using (var driver = new ChromeDriver(optionsChr))
                {
                    try
                    {
                        if (Ferramentas.LogarSistema(driver, xCliente))
                        {
                            driver.Navigate().GoToUrl(Settings.Default.Relatorios + relatorio.NumeroRelatorio);
                            driver.Navigate().Refresh();
                            Thread.Sleep(5000);

                            // Fecha o Popup
                            if (Ferramentas.ValidarElementoSeExiste(driver, "ilabspush-btn-close", "ClassName"))
                            {
                                driver.FindElementByClassName("ilabspush-btn-close").Click();
                            }

                            Ferramentas.LimparPastaDownload(Settings.Default.Download,
                                relatorio.NomeArquivoRelatorio);
                            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                            // Setar a data
                            var inputDateIn =
                                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                                    driver.FindElementByName("inicio")));
                            inputDateIn.Clear();

                            var dataInicial = new DateTime(2018, 1, 1);

                            //inputDateIn.SendKeys(dataInicial.ToString("   01012018"));
                            inputDateIn.SendKeys(xCliente.DataInicioRelatorio.Trim().Length == 7
                                ? dataInicial.ToString("   0" + xCliente.DataInicioRelatorio)
                                : dataInicial.ToString("   " + xCliente.DataInicioRelatorio));

                            Thread.Sleep(1000);

                            var inputDateEnd = driver.FindElementByName("fim");
                            inputDateEnd.Clear();
                            inputDateEnd.SendKeys("");
                            inputDateEnd.SendKeys("31");
                            inputDateEnd.SendKeys("12");
                            inputDateEnd.SendKeys(DateTime.Now.ToString("yyyy"));
                            Thread.Sleep(1000);

                            var searchDateButton =
                                driver.FindElementByXPath($"//*[@id='variaveis']/div[1]/span[{relatorio.NumeroSpan}]/a");
                            searchDateButton?.Click();

                            var waitTable = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                            waitTable.Until(
                                SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(
                                    By.ClassName("sorting_1")));

                            Thread.Sleep(10000);

                            // Baixar Relatorio   
                            var excelButton = driver.FindElementByClassName("buttons-html5");
                            excelButton.Click();

                            Thread.Sleep(5000);

                            // Move o relatório baixado para a pasta do respectivo cliente
                            Ferramentas.MoverRelatorioPasta(Settings.Default.CaminhoDestinoRelatorios,
                                xCliente.NomeCliente, relatorio.NomeArquivoRelatorio, relatorio.NumeroRelatorio);

                            // Atualiza data da ultima execução com sucesso do cliente e valida se quem chamou foi a rotina de erros
                            if (!ErrorDownload)
                            {
                                ServiceCliente.AtualizarCliente(xCliente.NomeCliente, false, true, false);
                            }
                        }
                        else
                        {
                            throw new ArgumentException("Erro ao logar usuário.");
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceCliente.AtualizarCliente(xCliente.NomeCliente, true, true, false);
                        Ferramentas.GravarLog("BaixarRelatorios - " + relatorio.NumeroRelatorio + " / Cliente: " + xCliente.NomeCliente, ex);
                        ErrorDownload = true;

                    }
                    finally
                    {
                        driver.Close();
                        driver.Quit();
                    }
                }
            }
        }
    }
}