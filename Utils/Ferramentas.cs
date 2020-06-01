using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using RoboBarbearia.Properties;

namespace RoboBarbearia.Utils
{
    public static class Ferramentas
    {
        private static readonly string ArquivoLogErro =
            Path.Combine(Settings.Default.CaminhoUsuarios, "ArquivoLogErro.txt");

        public static void LimparPastaDownload(string xCaminho, string xNomeArquivo)
        {
            try
            {
                var arquivoAntigo = new FileInfo(Path.Combine(xCaminho, xNomeArquivo));
                if (arquivoAntigo.Exists) arquivoAntigo.Delete();
            }
            catch (Exception ex)
            {
                GravarLog("LimparPastaDownload", ex);
            }
        }

        public static void MoverRelatorioPasta(string xPathCliente, string xNomeRelatorio,
            string xNomeArquivoRelatorio, string xNumeroRelatorio)
        {
            try
            {
                xPathCliente = xPathCliente + "Relatorio_" + xNumeroRelatorio;
                if (!Directory.Exists(xPathCliente)) Directory.CreateDirectory(xPathCliente);

                var arquivoRelatorioNovo =
                    new FileInfo(
                        Path.Combine(Settings.Default.Download, xNomeArquivoRelatorio));
                if (!arquivoRelatorioNovo.Exists) return;
                LimparPastaDownload(xPathCliente, xNomeRelatorio + ".xlsx");
                arquivoRelatorioNovo.MoveTo(xPathCliente + "\\" + xNomeRelatorio + ".xlsx");
            }
            catch (Exception ex)
            {
                GravarLog("MoverRelatorioPasta", ex);
            }
        }
        
        public static void MoverFinanceiroPasta(string xPathCliente, string xNomeCliente,
            string xNomeArquivo, string xNomeFinanceiro, string xContaBancaria, string xTipoData, string xTipoValor)
        {
            try
            {
                xPathCliente = xPathCliente + "Financeiro_" + xNomeFinanceiro + "_" + xContaBancaria + "_" + xTipoData + "_" + xTipoValor;
                if (!Directory.Exists(xPathCliente)) Directory.CreateDirectory(xPathCliente);

                var arquivoRelatorioNovo =
                    new FileInfo(
                        Path.Combine(Settings.Default.Download, xNomeArquivo));
                
                if (!arquivoRelatorioNovo.Exists) return;
                LimparPastaDownload(xPathCliente, xNomeCliente + ".xlsx");
                arquivoRelatorioNovo.MoveTo(xPathCliente + "\\" + xNomeCliente + ".xlsx");
            }
            catch (Exception ex)
            {
                GravarLog("MoverFinanceiroPasta", ex);
            }
        }

        public static void GravarLog(string xMsg, Exception xMensagemErro)
        {
            if (!File.Exists(ArquivoLogErro))
            {
                var arquivo = File.Create(ArquivoLogErro);
                arquivo.Close();
            }

            using (var textWriter = File.AppendText(ArquivoLogErro))
            {
                textWriter.Write("\r\nLog Entrada : ");
                textWriter.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
                textWriter.WriteLine("  :");
                textWriter.WriteLine($"  Erro rotina: {xMsg}");
                textWriter.WriteLine($"  :{xMensagemErro}");
                textWriter.WriteLine("------------------------------------");
            }
        }
        
        public static bool ValidarMaior365Dias(int xDia, int xMes, int xAno)
        {
            var data = new DateTime(xAno, xMes, xDia);
            
            var dataAtual = DateTime.Now;

            return ((data - dataAtual).Days < 365); 
        }

        public static bool ValidarElementoSeExiste(RemoteWebDriver xDriver, string xElement, string xTipoBusca)
        {
            var elementExiste = new List<IWebElement>();
            switch (xTipoBusca)
            {
                case "ClassName":
                    elementExiste.AddRange(xDriver.FindElementsByClassName(xElement));
                    break;
                case "XPath":
                    elementExiste.AddRange(xDriver.FindElementsByClassName(xElement));
                    break;
            }
            
            return elementExiste.Count > 0;
        }
    }
}