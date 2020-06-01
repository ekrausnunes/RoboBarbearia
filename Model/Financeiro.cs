using System;

namespace RoboBarbearia.Model
{
    public class Financeiro
    {
        public Financeiro(string nomeFinanceiro, string contasBancarias, string tipoData, string tipoValor,string nomeArquivo,
            bool ativoFinanceiro)
        {
            NomeFinanceiro = nomeFinanceiro ?? throw new ArgumentNullException(nameof(nomeFinanceiro));
            ContasBancarias = contasBancarias ?? throw new ArgumentNullException(nameof(contasBancarias));
            TipoData = tipoData ?? throw new ArgumentNullException(nameof(tipoData));
            TipoValor = tipoValor ?? throw new ArgumentNullException(nameof(tipoData));
            NomeArquivo = nomeArquivo ?? throw new ArgumentNullException(nameof(nomeArquivo));
            AtivoFinanceiro = ativoFinanceiro;
        }
        
        public string NomeFinanceiro { get; }
        public string ContasBancarias { get; }
        public string TipoData { get; }
        public string TipoValor { get; }
        public string NomeArquivo { get; }
        public bool AtivoFinanceiro { get; }
    }
}