using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.Extensions;
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
        private static List<Financeiro> _financeiroLista;

        private static void Main()
        {
            var nomeCliente = "";
            var rotinaRelatorio = false;
            var rotinaFinanceiro = false;
            try
            {
                BuscarRelatorios();
                BuscarClientes();
                BuscarFinanceiros();

                // Baixa os relatórios de todos os clientes que estão setados para baixar (gerarRelatorioCliente) e todos desatualizados (dataUltAtualizacao < dataAtual)
                foreach (var cliente in _clienteLista)
                {
                    nomeCliente = cliente.NomeCliente;
                    // Baixa todos os relatórios configurados do cliente
                    rotinaRelatorio = true;
                    if (cliente.GerarRelatorioCliente && cliente.DataUltAtualizacaoRelatorios != "ERRO" &&
                        cliente.DataUltAtualizacaoRelatorios != DateTime.Now.ToString("ddMMyy")) BaixarRelatorios(cliente);
                    rotinaRelatorio = false;
                    
                    // Baixa todos o financeiro configurados do cliente
                    rotinaFinanceiro = true;
                    if (cliente.GerarFinanceiroCliente && cliente.DataUltAtualizacaoFinanceiro != "ERRO" &&
                        cliente.DataUltAtualizacaoFinanceiro != DateTime.Now.ToString("ddMMyy")) BaixarFinanceiros(cliente);
                    rotinaFinanceiro = false;
                }

                // Baixa os relatórios de todos os clientes que estão setados para baixar (gerarRelatorioCliente) e que estão com ERRO (dataUltAtualizacao = ERRO)
                foreach (var cliente in _clienteLista)
                {
                    nomeCliente = cliente.NomeCliente;
                    if (cliente.GerarRelatorioCliente && cliente.DataUltAtualizacaoRelatorios == "ERRO")
                        BaixarRelatorios(cliente);
                    
                    if (cliente.GerarFinanceiroCliente && cliente.DataUltAtualizacaoFinanceiro == "ERRO")
                        BaixarFinanceiros(cliente);
                }
            }
            catch (Exception ex)
            {
                // Atualiza data da ultima execução com erro, para que seja executado novamente
                AtualizarCliente(nomeCliente, true, rotinaRelatorio, rotinaFinanceiro);
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
                                worksheet.Cells[row, 12].Value?.ToString(),
                                worksheet.Cells[row, 13].Value?.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 14].Value.ToString(),
                                worksheet.Cells[row, 15].Value?.ToString()
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
        
        private static void BuscarFinanceiros()
        {
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Financeiros.xlsx"));

                var workBook = package.Workbook;
                _financeiroLista = new List<Financeiro>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Financeiro"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Value.ToString()))
                            _financeiroLista.Add(new Financeiro(
                                worksheet.Cells[row, 1].Value.ToString().Trim(),
                                worksheet.Cells[row, 2].Value.ToString().Trim(),
                                worksheet.Cells[row, 3].Value.ToString().Trim(),
                                worksheet.Cells[row, 4].Value.ToString().Trim(),
                                worksheet.Cells[row, 5].Value.ToString().Trim(),
                                worksheet.Cells[row, 6].Value.ToString().Trim().ToUpper() == "SIM"));
                }

                package.Dispose();
            }
            catch (Exception ex)
            {
                Ferramentas.GravarLog("BuscarFinanceiros", ex);
            }
        }

        private static void AtualizarCliente(string pNomeCliente, bool xEhErro, bool xEhRelatorio, bool xEhFinanceiro)
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

                        if (!xEhFinanceiro) continue;
                        if (worksheet.Cells[row, 1].Value.ToString() == pNomeCliente)
                            worksheet.Cells[row, 15].Value = xEhErro ? "ERRO" : DateTime.Now.ToString("ddMMyy");
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

        private static void BaixarRelatorios(Cliente xCliente)
        {
            var optionsChr = new ChromeOptions();
            // Roda sem o browser
            //optionsChr.AddArgument("--headless");
            optionsChr.AddArgument("--ignore-certificate-errors");
            optionsChr.AddArgument("--ignore-ssl-errors");

            // Inicializa o Chrome Driver
            using (var driver = new ChromeDriver(optionsChr))
            {
                if (LogarSistema(driver, xCliente))
                {
                    try
                    {
                        foreach (var relatorio in _relatorioLista.Where(relatorio => relatorio.AtivoRelatorio))
                        {
                            driver.Navigate().GoToUrl(Settings.Default.Relatorios +
                                                      relatorio.NumeroRelatorio);
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
                                wait.Until(ExpectedConditions.ElementToBeClickable(
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
                        }

                        // Atualiza data da ultima execução com sucesso do cliente
                        AtualizarCliente(xCliente.NomeCliente, false, true, false);
                    }
                    catch (Exception ex)
                    {
                        Ferramentas.GravarLog("BaixarRelatorios", ex);
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
        
        private static void BaixarFinanceiros(Cliente xCliente)
        {
            var optionsChr = new ChromeOptions();
            optionsChr.AddArgument("--ignore-certificate-errors");
            optionsChr.AddArgument("--ignore-ssl-errors");


            // Inicializa o Chrome Driver
            using (var driver = new ChromeDriver(optionsChr))
            {
                if (LogarSistema(driver, xCliente))
                {
                    try
                    {
                        driver.Navigate().GoToUrl(Settings.Default.Financeiro);
                        
                        foreach (var financeiro in _financeiroLista.Where(financeiro => financeiro.AtivoFinanceiro))
                        {
                            driver.Navigate().Refresh();
                            Thread.Sleep(5000);
                            
                            // Fecha o Popup
                            if (Ferramentas.ValidarElementoSeExiste(driver, "ilabspush-btn-close", "ClassName"))
                            {
                                driver.FindElementByClassName("ilabspush-btn-close").Click();
                            }
                            
                            Ferramentas.LimparPastaDownload(Settings.Default.Download,
                                financeiro.NomeArquivo);
                            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                            // Setar a aba do financeiro
                            var tabFinan = driver.FindElementById("extrato");
                            switch (financeiro.NomeFinanceiro.ToUpper())
                            {
                                case "EXTRATO":
                                    tabFinan = driver.FindElementById("extrato");
                                    break;
                                case "DESPESAS":
                                    tabFinan = driver.FindElementById("contasApagar");
                                    break;
                                case "RECEITAS":
                                    tabFinan = driver.FindElementById("contasAreceber");
                                    break;
                            }
                            // As vezes aparece um popup em cima do elemento e com isso é necessário forçar uma execução em JS.
                            driver.ExecuteJavaScript("arguments[0].click();", tabFinan);

                            // Setar a data
                            var inputDateIn =
                                wait.Until(ExpectedConditions.ElementToBeClickable(
                                    driver.FindElementById("dataini")));
                            inputDateIn.Click();
                            inputDateIn.Clear();
                            
                            var dataInitial = new DateTime(2018, 1, 1);
                            var dateInitialUser = (xCliente.DataInicioFinanceiro.Trim().Length == 7
                                ? dataInitial.ToString("0" + xCliente.DataInicioFinanceiro)
                                : dataInitial.ToString(xCliente.DataInicioFinanceiro));

                            // O site valida se a data inicial e a data final tem mais de 365 dias de <>, caso tenha eu tiro 365 da data atual e não pego a data do usuário
                            inputDateIn.SendKeys(Ferramentas.ValidarMaior365Dias(int.Parse(dateInitialUser.Substring(0, 2)),
                                int.Parse(dateInitialUser.Substring(2, 2)),
                                int.Parse(dateInitialUser.Substring(4, 4)))
                                ? (DateTime.Now.AddDays(-365)).ToString( "    ddMMyyyy")
                                : dateInitialUser);

                            Thread.Sleep(1000);

                            // Setar Contas Bancarias
                            switch (financeiro.ContasBancarias.ToUpper())
                            {
                                case "TODOS":
                                    driver.FindElementByXPath("//*[@id='conta_bancaria']/option[1]").Click();
                                    break;
                                case "CAIXA":
                                    driver.FindElementByXPath("//*[@id='conta_bancaria']/option[2]").Click();
                                    break;
                                case "AVECPASS":
                                    // Caso não tenha esta opção, vai para o próximo.
                                    if (Ferramentas.ValidarElementoSeExiste(driver,
                                        "//*[@id='conta_bancaria']/option[3]", "XPath"))
                                    {
                                        driver.FindElementByXPath("//*[@id='conta_bancaria']/option[3]").Click(); 
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                    break;
                            }

                            // Setar tipo de conta
                            switch (financeiro.TipoData.ToUpper())
                            {
                                case "QUITAÇÃO":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[1]").Click();
                                    break;
                                case "COMPETÊNCIA":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[2]").Click();
                                    break;
                                case "VENCIMENTO":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[3]").Click();
                                    break;
                            }
                            
                            // Selecionar Bruto/Líquido
                            switch (financeiro.TipoValor.ToUpper())
                            {
                                case "BRUTO":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[7]/label[1]").Click();
                                    break;
                                case "LÍQUIDO":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[7]/label[2]").Click();
                                    break;
                            }

                            // Executar a busca
                            driver.FindElementByXPath("//*[@id='divExtrato']/form/div[2]/a").Click();

                            var waitTable = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                            waitTable.Until(
                                ExpectedConditions.VisibilityOfAllElementsLocatedBy(
                                    By.ClassName("sorting_1")));

                            Thread.Sleep(10000);
                            
                            // Baixar Relatorio   
                            driver.FindElementByClassName("buttons-html5").Click();

                            Thread.Sleep(5000);

                            // Move o relatório baixado para a pasta do respectivo cliente
                            Ferramentas.MoverFinanceiroPasta(Settings.Default.CaminhoDestinoFinanceiros,
                                xCliente.NomeCliente, financeiro.NomeArquivo, financeiro.NomeFinanceiro, 
                                financeiro.ContasBancarias, financeiro.TipoData, financeiro.TipoValor);
                        }
                        // Atualiza data da ultima execução com sucesso do cliente
                        AtualizarCliente(xCliente.NomeCliente, false, false, true);
                    }
                    catch (Exception ex)
                    {
                        Ferramentas.GravarLog("BaixarFinanceiros", ex);
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