using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RoboBarbearia.Model
{
    public class Cliente
    {
        public Cliente() {}

        public string nomeCliente { get; set; }
        public string usuario { get; set; }
        public bool gerarRelatorioCliente { get; set; }
        public string admSalaoVIPCliente { get; set; }
        public string loginSite { get; set; }
        public string senhaSite { get; set; }
        public string donoCliente { get; set; }
        public string dataInicio { get; set; }

        public Cliente(string nomeCliente, string usuario, bool gerarRelatorioCliente, string admSalaoVIPCliente, string loginSite, string senhaSite, string donoCliente, string dataInicio)
        {
            this.nomeCliente = nomeCliente;
            this.usuario = usuario;
            this.gerarRelatorioCliente = gerarRelatorioCliente;
            this.admSalaoVIPCliente = admSalaoVIPCliente;
            this.loginSite = loginSite;
            this.senhaSite = senhaSite;
            this.donoCliente = donoCliente;
            this.dataInicio = dataInicio;
        }
    }
}
