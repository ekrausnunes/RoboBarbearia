using System;
using System.IO;
using RoboBarbearia.Properties;

namespace RoboBarbearia.Utils
{
    public static class Ferramentas
    {
        private static readonly string ArquivoLogErro =
            Path.Combine(Settings.Default.CaminhoUsuarios, "ArquivoLogErro.txt");

        public static void LimparPastaDownload(string xCaminho, string xNomeArquivoRelatorio)
        {
            try
            {
                var arquivoRelatorioAntigo = new FileInfo(Path.Combine(xCaminho, xNomeArquivoRelatorio));
                if (arquivoRelatorioAntigo.Exists) arquivoRelatorioAntigo.Delete();
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
    }
}