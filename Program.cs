using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using RoboBarbearia.Model;
using RoboBarbearia.Properties;
using RoboBarbearia.Utils;
using ExpectedConditions = SeleniumExtras.WaitHelpers.ExpectedConditions;

namespace RoboBarbearia
{
    internal static class Program
    {
        private static List<Cliente> _clienteLista;
        private static List<Relatorio> _relatorioLista;

        private static void Main()
        {
            var nomeCliente = "";
            try
            {
                BuscarRelatorios();
                BuscarClientes();

                // Baixa os relatórios de todos os clientes que estão setados para baixar (gerarRelatorioCliente) e todos desatualizados (dataUltAtualizacao < dataAtual)
                foreach (var cliente in _clienteLista)
                {
                    nomeCliente = cliente.NomeCliente;
                    if (cliente.GerarRelatorioCliente && cliente.DataUltAtualizacao != "ERRO" &&
                        cliente.DataUltAtualizacao != DateTime.Now.ToString("ddMMyy")) BaixarRelatorios(cliente);
                }

                // Baixa os relatórios de todos os clientes que estão setados para baixar (gerarRelatorioCliente) e que estão com ERRO (dataUltAtualizacao = ERRO)
                foreach (var cliente in _clienteLista)
                {
                    nomeCliente = cliente.NomeCliente;
                    if (cliente.GerarRelatorioCliente && cliente.DataUltAtualizacao == "ERRO")
                        BaixarRelatorios(cliente);
                }
            }
            catch (Exception ex)
            {
                // Atualiza data da ultima execução com erro, para que seja executado novamente
                AtualizarCliente(nomeCliente, true);
                Ferramentas.GravarLog("Principal", ex);
            }
        }

        private static bool LogarSistema(RemoteWebDriver xDriver, Cliente pCliente)
        {
            try
            {
                // Vai para pagina Login do site
                xDriver.Navigate().GoToUrl(pCliente.AdmSalaoVipCliente);

                // Pega o elemento Login/Senha
                var waitLogin = new WebDriverWait(xDriver, TimeSpan.FromSeconds(60));
                waitLogin.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.ClassName("container-form")));

                // Fecha o Popup
                //IWebElement closePopup = xDriver.FindElementByClassName("ilabspush-btn-close");
                //closePopup.Click();

                var userNameField = xDriver.FindElementById("formEmail");
                var userPasswordField = xDriver.FindElementById("formSenha");

                // Pega a classe btn-login, botão login
                var loginButton = xDriver.FindElementByClassName("btn-login");

                // Passa Login/Senha
                if (pCliente.DonoCliente?.Trim().ToUpper() == Settings.Default.Admin)
                {
                    userNameField.SendKeys(Settings.Default.Login);
                    userPasswordField.SendKeys(Settings.Default.Senha);
                }
                else
                {
                    userNameField.SendKeys(pCliente.LoginSite);
                    userPasswordField.SendKeys(pCliente.SenhaSite);
                }

                loginButton.Click();

                Thread.Sleep(5000);
                return true;
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("LogarSistema / Cliente: " + pCliente.NomeCliente, ex);
                return false;
            }
        }

        private static void BuscarClientes()
        {
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Usuarios.xlsx"));
                var workBook = package.Workbook;
                _clienteLista = new List<Cliente>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Planilha1"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 4].Value.ToString()))
                            _clienteLista.Add(new Cliente(
                                worksheet.Cells[row, 1].Value?.ToString().Trim(),
                                worksheet.Cells[row, 5].Value?.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 6].Value?.ToString().Trim(),
                                worksheet.Cells[row, 7].Value?.ToString().Trim(),
                                worksheet.Cells[row, 8].Value?.ToString().Trim(),
                                worksheet.Cells[row, 9].Value?.ToString().Trim(),
                                worksheet.Cells[row, 10].Value.ToString(),
                                worksheet.Cells[row, 12].Value?.ToString()
                            ));
                }

                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarClientes", ex);
            }
        }

        private static void BuscarRelatorios()
        {
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Relatorios.xlsx"));

                var workBook = package.Workbook;
                _relatorioLista = new List<Relatorio>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Relatorios"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 2].Value.ToString()))
                            _relatorioLista.Add(new Relatorio(worksheet.Cells[row, 2].Value.ToString().Trim(),
                                worksheet.Cells[row, 3].Value.ToString().Trim(),
                                worksheet.Cells[row, 4].Value.ToString().Trim().ToUpper() == "SIM"));
                }

                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarRelatorios", ex);
            }
        }

        private static void AtualizarCliente(string pNomeCliente, bool xEhErro)
        {
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Usuarios.xlsx"));
                var workBook = package.Workbook;
                _clienteLista = new List<Cliente>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Planilha1"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (worksheet.Cells[row, 1].Value.ToString() == pNomeCliente)
                            worksheet.Cells[row, 12].Value = xEhErro ? "ERRO" : DateTime.Now.ToString("ddMMyy");
                }

                package.Save();
                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarClientes", ex);
            }
        }

        private static void BaixarRelatorios(Cliente xCliente)
        {
            var optionsChr = new ChromeOptions();
            // Roda sem o browser
            //optionsChr.AddArgument("--headless");
            optionsChr.AddArgument("--ignore-certificate-errors");
            optionsChr.AddArgument("--ignore-ssl-errors");
            //optionsChr.AddArgument("enable-automation"); 
            //optionsChr.AddArgument("--no-sandbox");
            //optionsChr.AddArgument("--disable-infobars");
            //optionsChr.AddArgument("--disable-dev-shm-usage");
            //optionsChr.AddArgument("--disable-browser-side-navigation");
            //optionsChr.AddArgument("--disable-gpu");

            // Inicializa o Chrome Driver
            using (var driver = new ChromeDriver(optionsChr))
            {
                if (LogarSistema(driver, xCliente))
                {
                    foreach (var relatorio in _relatorioLista.Where(relatorio => relatorio.AtivoRelatorio))
                    {
                        driver.Navigate().GoToUrl(Settings.Default.Relatorios +
                                                  relatorio.NumeroRelatorio);
                        driver.Navigate().Refresh();

                        Thread.Sleep(5000);

                        Ferramentas.LimparPastaDownload(Settings.Default.Download,
                            relatorio.NomeArquivoRelatorio);
                        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                        // Setar a data
                        var inputDateIn =
                            wait.Until(ExpectedConditions.ElementToBeClickable(
                                driver.FindElementByName("inicio")));
                        inputDateIn.Clear();

                        var dataInicial = new DateTime(2018, 1, 1);

                        //inputDateIn.SendKeys(dataInicial.ToString("   01012018"));
                        inputDateIn.SendKeys(xCliente.DataInicio.Trim().Length == 7
                            ? dataInicial.ToString("   0" + xCliente.DataInicio)
                            : dataInicial.ToString("   " + xCliente.DataInicio));

                        Thread.Sleep(1000);

                        var inputDateEnd = driver.FindElementByName("fim");
                        inputDateEnd.Clear();
                        inputDateEnd.SendKeys("");
                        inputDateEnd.SendKeys("31");
                        inputDateEnd.SendKeys("12");
                        inputDateEnd.SendKeys(DateTime.Now.ToString("yyyy"));
                        Thread.Sleep(1000);

                        var searchDateButton =
                            driver.FindElementByXPath("//*[@id='variaveis']/span[3]/a");
                        searchDateButton?.Click();

                        var waitTable = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                        waitTable.Until(
                            ExpectedConditions.VisibilityOfAllElementsLocatedBy(
                                By.ClassName("sorting_1")));

                        Thread.Sleep(10000);
                        // Baixar Relatorio   
                        var excelButton = driver.FindElementByClassName("buttons-html5");
                        excelButton.Click();

                        Thread.Sleep(5000);

                        // Move o relatório baixado para a pasta do respectivo cliente
                        Ferramentas.MoverRelatorioPasta(Settings.Default.CaminhoDestinoRelatorios,
                            xCliente.NomeCliente, relatorio.NomeArquivoRelatorio, relatorio.NumeroRelatorio);

                        // Atualiza data da ultima execução com sucesso do cliente
                        AtualizarCliente(xCliente.NomeCliente, false);
                    }
                }
                else
                {
                    driver.Quit();
                    throw new ArgumentException("Erro ao logar usuário.");
                }

                driver.Quit();
            }
        }
    }
}