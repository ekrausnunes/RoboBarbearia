using System;
using System.Collections.Generic;
using System.Linq;
using RoboBarbearia.Model;
using RoboBarbearia.Service;

namespace RoboBarbearia
{
    internal static class Program
    {
        private static List<Cliente> _clienteLista;
        private static List<Relatorio> _relatorioLista;
        private static List<Financeiro> _financeiroLista;

        private static void Main()
        {
            _relatorioLista = ServiceRelatorio.BuscarRelatorios();
            _clienteLista = ServiceCliente.BuscarClientes();
            _financeiroLista = ServiceFinanceiro.BuscarFinanceiros();
            
            // !!!!!Cada um baixa separado pois o consumo de memória do chrome aumenta muito!!!!!
            
            // Baixa o relatórios de todos os clientes que estão setados para baixar (gerarFinanceiroCliente) e todos desatualizados (dataUltAtualizacao < dataAtual)
            foreach (var cliente in _clienteLista.Where(cliente => cliente.GerarRelatorioCliente && cliente.DataUltAtualizacaoRelatorios != "ERRO" &&
                                                                   cliente.DataUltAtualizacaoRelatorios != DateTime.Now.ToString("ddMMyy")))
            {
                ServiceRelatorio.BaixarRelatorios(cliente, _relatorioLista);
            }

            // Atualiza a lista com casos de relatorios com erros
            _clienteLista.Clear();
            _clienteLista = ServiceCliente.BuscarClientes();

            // Baixa os relatórios de todos os clientes que estão setados para baixar (gerarRelatorioCliente) e que estão com ERRO (dataUltAtualizacao = ERRO)
            foreach (var cliente in _clienteLista.Where(cliente => cliente.GerarRelatorioCliente && cliente.DataUltAtualizacaoRelatorios == "ERRO"))
            {
                ServiceRelatorio.BaixarRelatorios(cliente, _relatorioLista);
            }
            
            // Baixa o financeiro de todos os clientes que estão setados para baixar (gerarFinanceiroCliente) e todos desatualizados (dataUltAtualizacao < dataAtual)
            foreach (var cliente in _clienteLista.Where(cliente => cliente.GerarFinanceiroCliente && cliente.DataUltAtualizacaoFinanceiro != "ERRO" &&
                                                                   cliente.DataUltAtualizacaoFinanceiro != DateTime.Now.ToString("ddMMyy")))
            {
                ServiceFinanceiro.BaixarFinanceiros(cliente, _financeiroLista);
            }
            
            // Atualiza a lista com casos de relatorios com erros
            _clienteLista.Clear();
            _clienteLista = ServiceCliente.BuscarClientes();

            // Baixa o financeiro de todos os clientes que estão setados para baixar (gerarFinanceiroCliente) e que estão com ERRO (dataUltAtualizacao = ERRO)
            foreach (var cliente in _clienteLista.Where(cliente => cliente.GerarFinanceiroCliente && cliente.DataUltAtualizacaoFinanceiro == "ERRO"))
            {
                ServiceFinanceiro.BaixarFinanceiros(cliente, _financeiroLista);
            }
        }
    }
}