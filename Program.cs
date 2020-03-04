using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using RoboBarbearia.Model;
using System.Collections.Generic;

namespace WebDriverTest
{
    class Program
    {
        private static readonly string arquivoLogErro = Path.Combine(RoboBarbearia.Properties.Settings.Default.CaminhoUsuarios, "ArquivoLogErro.txt");
        private static List<Cliente> clienteLista;
        private static List<Relatorio> relatorioLista;

        static void Main(string[] args)
        {
            try
            {                
                ChromeOptions optionsChr = new ChromeOptions();
                // Roda sem o browser
                //optionsChr.AddArgument("--headless");

                // Inicializa o Chrome Driver
                using (ChromeDriver driver = new ChromeDriver(optionsChr))                
                {
                    BuscarRelatorios();
                    BuscarClientes();
     
                    foreach (Cliente xCliente in clienteLista)
                    {
                        if (xCliente.gerarRelatorioCliente)
                        {
                            if (LogarSistema(driver, xCliente))
                            {
                                foreach (Relatorio xRelatorio in relatorioLista)
                                {
                                    if (xRelatorio.ativoRelatorio)
                                    {
                                        driver.Navigate().GoToUrl(RoboBarbearia.Properties.Settings.Default.Relatorios + xRelatorio.numeroRelatorio);
                                        driver.Navigate().Refresh();
                                        
                                        LimparPastaDownload(RoboBarbearia.Properties.Settings.Default.Download, xRelatorio.nomeArquivoRelatorio);
                                        System.Threading.Thread.Sleep(5000);
                                        // Setar a data
                                        IWebElement inputDateIn = driver.FindElementByName("inicio");
                                        inputDateIn.Clear();

                                        DateTime dataInicial = new DateTime(2018, 1, 1); 
                                        
                                        //inputDateIn.SendKeys(dataInicial.ToString("   01012018"));
                                        if (xCliente.dataInicio.Trim().Length == 7) {
                                            inputDateIn.SendKeys(dataInicial.ToString("   0" + xCliente.dataInicio));
                                        } else {
                                            inputDateIn.SendKeys(dataInicial.ToString("   " + xCliente.dataInicio));
                                        }
                                        
                                        System.Threading.Thread.Sleep(1000);

                                        IWebElement inputDateEnd = driver.FindElementByName("fim");
                                        inputDateEnd.Clear();
                                        inputDateEnd.SendKeys("");
                                        inputDateEnd.SendKeys("31");
                                        inputDateEnd.SendKeys("12");
                                        inputDateEnd.SendKeys(DateTime.Now.ToString("yyyy"));
                                        System.Threading.Thread.Sleep(1000);

                                        IWebElement searchDateButton = driver.FindElementByXPath("//*[@id='variaveis']/span[3]/a");
                                        searchDateButton.Click();

                                        System.Threading.Thread.Sleep(300000);

                                        // Baixar Relatorio                                   
                                        IWebElement excelButton = driver.FindElementByClassName("buttons-html5");
                                        excelButton.Click();

                                        System.Threading.Thread.Sleep(1000);

                                        // Move o relatório baixado para a pasta do respectivo cliente
                                        MoverRelatorioPasta(RoboBarbearia.Properties.Settings.Default.CaminhoDestinoRelatorios, xCliente.nomeCliente, xRelatorio.nomeArquivoRelatorio, xRelatorio.numeroRelatorio);
                                    }
                                }
                            }
                            else
                            {
                                driver.Quit();
                                throw new ArgumentException("Erro ao logar usuário.");
                            }
                        }                     
                    }               
                    driver.Quit();
                }                
            }
            catch (Exception ex)
            {
                GravarLog("Principal / Baixar Relatório", ex);                            
            }
        }

        private static bool LogarSistema(ChromeDriver xDriver, Cliente pCliente)
        {
            try
            {                
                // Vai para pagina Login do site
                xDriver.Navigate().GoToUrl(pCliente.admSalaoVIPCliente);
                xDriver.Navigate().Refresh();

                System.Threading.Thread.Sleep(2000);

                // Pega o elemento Login/Senha
                IWebElement userNameField = xDriver.FindElementById("formEmail");
                IWebElement userPasswordField = xDriver.FindElementById("formSenha");

                // Pega a classe btn-login, botão login
                IWebElement loginButton = xDriver.FindElementByClassName("btn-login");

                // Passa Login/Senha
                if (pCliente.donoCliente?.Trim().ToUpper() == "RODRIGO") {
                    userNameField.SendKeys(RoboBarbearia.Properties.Settings.Default.Login);
                    userPasswordField.SendKeys(RoboBarbearia.Properties.Settings.Default.Senha);
                } else
                {
                    userNameField.SendKeys(pCliente.loginSite);
                    userPasswordField.SendKeys(pCliente.senhaSite);
                }

                loginButton.Click();

                System.Threading.Thread.Sleep(5000);

                IWebElement nomeCliente = xDriver.FindElementByClassName("titulo-header");

                return true;
            }
            catch (Exception ex)
            {
                GravarLog("LogarSistema / Cliente: " + pCliente.nomeCliente, ex);
                return false;
            }
        }

        private static void BuscarClientes()
        {
            try
            {
                ExcelPackage package = new ExcelPackage(new FileInfo(RoboBarbearia.Properties.Settings.Default.CaminhoUsuarios + "Usuarios.xlsx"));
                var workBook = package.Workbook;
                clienteLista = new List<Cliente>();

                if (workBook != null)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Planilha1"];
                    int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 4].Value.ToString()))
                        {
                            clienteLista.Add(new Cliente(
                                worksheet.Cells[row, 1].Value?.ToString().Trim(),
                                worksheet.Cells[row, 2].Value?.ToString().Trim(),
                                worksheet.Cells[row, 5].Value?.ToString().Trim().ToUpper() == "SIM",
                                worksheet.Cells[row, 6].Value?.ToString().Trim(),
                                worksheet.Cells[row, 7].Value?.ToString().Trim(),
                                worksheet.Cells[row, 8].Value?.ToString().Trim(),
                                worksheet.Cells[row, 9].Value?.ToString().Trim(),
                                worksheet.Cells[row, 10].Value.ToString()
                                ));
                        }
                    }
                }
                package.Dispose();
            }
            catch (Exception ex)
            {
                GravarLog("BuscarClientes", ex);
            }
        }

        private static void BuscarRelatorios()
        {
            try
            {
                ExcelPackage package = new ExcelPackage(new FileInfo(RoboBarbearia.Properties.Settings.Default.CaminhoUsuarios + "Relatorios.xlsx"));

                var workBook = package.Workbook;
                relatorioLista = new List<Relatorio>();

                if (workBook != null)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Relatorios"];
                    int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;

                    // Começa depois do cabeçalho
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 2].Value.ToString()))
                        {
                            relatorioLista.Add(new Relatorio(
                                worksheet.Cells[row, 1].Value.ToString().Trim(),
                                worksheet.Cells[row, 2].Value.ToString().Trim(),
                                worksheet.Cells[row, 3].Value.ToString().Trim(),
                                worksheet.Cells[row, 4].Value.ToString().Trim().ToUpper() == "SIM"));
                        }
                    }
                }
                package.Dispose();
            }
            catch (Exception ex)
            {
                GravarLog("BuscarRelatorios", ex);
            }
        }

        private static void LimparPastaDownload(string xCaminho, string xNomeArquivoRelatorio)
        {
            try
            {
                FileInfo arquivoRelatorioAntigo = new FileInfo(Path.Combine(xCaminho, xNomeArquivoRelatorio));
                if (arquivoRelatorioAntigo.Exists)
                {
                    arquivoRelatorioAntigo.Delete();
                }
            }
            catch (Exception ex)
            {
                GravarLog("LimparPastaDownload", ex);
            }
        }

        private static void MoverRelatorioPasta(string xPathCliente, string xNomeRelatorio, string xNomeArquivoRelatorio, string xNumeroRelatorio)
        {
            try
            {
                xPathCliente = xPathCliente + "Relatorio_" + xNumeroRelatorio;
                if (!Directory.Exists(xPathCliente))
                {
                    Directory.CreateDirectory(xPathCliente);
                }

                FileInfo arquivoRelatoriNovo = new FileInfo(Path.Combine(RoboBarbearia.Properties.Settings.Default.Download, xNomeArquivoRelatorio));
                if (arquivoRelatoriNovo.Exists)
                {
                    LimparPastaDownload(xPathCliente, xNomeRelatorio + ".xlsx");
                    arquivoRelatoriNovo.MoveTo(xPathCliente + "\\" + xNomeRelatorio + ".xlsx");
                }
            }
            catch (Exception ex)
            {
                GravarLog("MoverRelatorioPasta", ex);
            }
        }

        private static void GravarLog(string xMsg, Exception xMensagemErro)
        {
            try
            {
                if (!File.Exists(arquivoLogErro))
                {
                    FileStream arquivo = File.Create(arquivoLogErro);
                    arquivo.Close();
                }

                using (StreamWriter textWriter = File.AppendText(arquivoLogErro))
                {
                    textWriter.Write("\r\nLog Entrada : ");
                    textWriter.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
                    textWriter.WriteLine("  :");
                    textWriter.WriteLine($"  Erro rotina: {xMsg}");
                    textWriter.WriteLine($"  :{xMensagemErro}");
                    textWriter.WriteLine("------------------------------------");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
