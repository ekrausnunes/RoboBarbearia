namespace RoboBarbearia.Model
{
    public class Cliente
    {
        public Cliente(string nomeCliente, bool gerarRelatorioCliente, string admSalaoVipCliente,
            string loginSite, string senhaSite, string donoCliente, string dataInicio, string dataUltAtualizacao)
        {
            NomeCliente = nomeCliente;
            GerarRelatorioCliente = gerarRelatorioCliente;
            AdmSalaoVipCliente = admSalaoVipCliente;
            LoginSite = loginSite;
            SenhaSite = senhaSite;
            DonoCliente = donoCliente;
            DataInicio = dataInicio;
            DataUltAtualizacao = dataUltAtualizacao;
        }

        public string NomeCliente { get; }
        public bool GerarRelatorioCliente { get; }
        public string AdmSalaoVipCliente { get; }
        public string LoginSite { get; }
        public string SenhaSite { get; }
        public string DonoCliente { get; }
        public string DataInicio { get; }
        public string DataUltAtualizacao { get; }
    }
}