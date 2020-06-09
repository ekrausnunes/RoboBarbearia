using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using RoboBarbearia.Model;
using RoboBarbearia.Properties;
using RoboBarbearia.Utils;

namespace RoboBarbearia.Service
{
    public static class ServiceFinanceiro
    {
        public static List<Financeiro> BuscarFinanceiros()
        {
            List<Financeiro> financeiroLista = null;
            try
            {
                var package =
                    new ExcelPackage(new FileInfo(Settings.Default.CaminhoUsuarios +
                                                  "Financeiros.xlsx"));

                var workBook = package.Workbook;
                financeiroLista = new List<Financeiro>();

                if (workBook != null)
                {
                    var worksheet = package.Workbook.Worksheets["Financeiro"];
                    var rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (var row = 2; row <= rowCount; row++)
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Value.ToString()))
                            financeiroLista.Add(new Financeiro(
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
            return new List<Financeiro>(financeiroLista ?? throw new InvalidOperationException("Rotina BuscarFinanceiros, retornou null!"));
        }

        public static void BaixarFinanceiros(Cliente xCliente, IEnumerable<Financeiro> xFinanceiroLista)
        {
            foreach (var financeiro in xFinanceiroLista.Where(financeiro => financeiro.AtivoFinanceiro))
            {
                var optionsChr = new ChromeOptions();
                optionsChr.AddArgument("--ignore-certificate-errors");
                optionsChr.AddArgument("--ignore-ssl-errors");

                // Inicializa o Chrome Driver
                using (var driver = new ChromeDriver(optionsChr))
                {
                    if (Ferramentas.LogarSistema(driver, xCliente))
                    {
                        try
                        {
                            driver.Navigate().GoToUrl(Settings.Default.Financeiro);
                            
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
                                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                                    driver.FindElementById("dataini")));
                            inputDateIn.Click();
                            inputDateIn.Clear();

                            var dataInitial = new DateTime(2018, 1, 1);
                            var dateInitialUser = (xCliente.DataInicioFinanceiro.Trim().Length == 7
                                ? dataInitial.ToString("0" + xCliente.DataInicioFinanceiro)
                                : dataInitial.ToString(xCliente.DataInicioFinanceiro));

                            // O site valida se a data inicial e a data final tem mais de 365 dias de <>, caso tenha eu tiro 365 da data atual e não pego a data do usuário
                            inputDateIn.SendKeys(Ferramentas.ValidarMaior365Dias(
                                int.Parse(dateInitialUser.Substring(0, 2)),
                                int.Parse(dateInitialUser.Substring(2, 2)),
                                int.Parse(dateInitialUser.Substring(4, 4)))
                                ? (DateTime.Now.AddDays(-365)).ToString("    ddMMyyyy")
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
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[1]")
                                        .Click();
                                    break;
                                case "COMPETÊNCIA":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[2]")
                                        .Click();
                                    break;
                                case "VENCIMENTO":
                                    driver.FindElementByXPath("//*[@id='campoRefinarBusca']/label[4]/select/option[3]")
                                        .Click();
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
                                SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(
                                    By.ClassName("sorting_1")));

                            Thread.Sleep(10000);

                            // Baixar Relatorio   
                            driver.FindElementByClassName("buttons-html5").Click();

                            Thread.Sleep(5000);

                            // Move o relatório baixado para a pasta do respectivo cliente
                            Ferramentas.MoverFinanceiroPasta(Settings.Default.CaminhoDestinoFinanceiros,
                                xCliente.NomeCliente, financeiro.NomeArquivo, financeiro.NomeFinanceiro,
                                financeiro.ContasBancarias, financeiro.TipoData, financeiro.TipoValor);

                            // Atualiza data da ultima execução com sucesso do cliente
                            ServiceCliente.AtualizarCliente(xCliente.NomeCliente, false, false, true);
                        }
                        catch (Exception ex)
                        {
                            ServiceCliente.AtualizarCliente(xCliente.NomeCliente, true, false, true);
                            Ferramentas.GravarLog("BaixarFinanceiros", ex);
                        }
                    }
                    else
                    {
                        throw new ArgumentException("Erro ao logar usuário.");
                    }

                    driver.Close();
                    driver.Quit();
                }
            }
        }
    }
}