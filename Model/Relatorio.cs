﻿using System;

namespace RoboBarbearia.Model
{
    public class Relatorio
    {
        public Relatorio(string numeroRelatorio, string nomeArquivoRelatorio, bool ativoRelatorio, string numeroSpan)
        {
            NumeroRelatorio = numeroRelatorio ?? throw new ArgumentNullException(nameof(numeroRelatorio));
            NomeArquivoRelatorio =
                nomeArquivoRelatorio ?? throw new ArgumentNullException(nameof(nomeArquivoRelatorio));
            AtivoRelatorio = ativoRelatorio;
            NumeroSpan = numeroSpan;
        }

        public string NumeroRelatorio { get; }
        public string NomeArquivoRelatorio { get; }
        public bool AtivoRelatorio { get; }
        public string NumeroSpan { get; }
    }
}