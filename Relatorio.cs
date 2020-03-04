using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RoboBarbearia.Model
{
    public class Relatorio
    {
        public Relatorio() {}

        public string nomeRelatorio { get; set; }
        public string numeroRelatorio { get; set; }
        public string nomeArquivoRelatorio { get; set; }
        public bool ativoRelatorio { get; set; }

        public Relatorio(string nomeRelatorio, string numeroRelatorio, string nomeArquivoRelatorio, bool ativoRelatorio)
        {
            this.nomeRelatorio = nomeRelatorio;
            this.numeroRelatorio = numeroRelatorio;
            this.nomeArquivoRelatorio = nomeArquivoRelatorio;
            this.ativoRelatorio = ativoRelatorio;
        }
    }
}
