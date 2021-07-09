namespace RoboBarbearia.Model
{
    public class Cliente
    {
        public Cliente(string nomeCliente, bool gerarRelatorioCliente, string admSalaoVipCliente,
            string loginSite, string senhaSite, string donoCliente, string dataInicioRelatorio, string dataUltAtualizacaoRelatorios, 
            bool gerarFinanceiroCliente, string dataInicioFinanceiro, string dataUltAtualizacaoFinanceiro)
        {
            NomeCliente = nomeCliente;
            GerarRelatorioCliente = gerarRelatorioCliente;
            AdmSalaoVipCliente = admSalaoVipCliente;
            LoginSite = loginSite;
            SenhaSite = senhaSite;
            DonoCliente = donoCliente;
            DataInicioRelatorio = dataInicioRelatorio;
            DataUltAtualizacaoRelatorios = dataUltAtualizacaoRelatorios;
            GerarFinanceiroCliente = gerarFinanceiroCliente;
            DataInicioFinanceiro = dataInicioFinanceiro;
            DataUltAtualizacaoFinanceiro = dataUltAtualizacaoFinanceiro;
        }

        public string NomeCliente { get; }
        public bool GerarRelatorioCliente { get; }
        public string AdmSalaoVipCliente { get; }
        public string LoginSite { get; }
        public string SenhaSite { get; }
        public string DonoCliente { get; }
        public string DataInicioRelatorio { get; }
        public string DataUltAtualizacaoRelatorios { get; }
        public bool GerarFinanceiroCliente { get; }
        public string DataInicioFinanceiro { get; }
        public string DataUltAtualizacaoFinanceiro { get; }
    }
}