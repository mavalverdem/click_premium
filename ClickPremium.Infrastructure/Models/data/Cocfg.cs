using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Parametros de configuracion
/// </summary>
public partial class Cocfg
{
    public string Codemp { get; set; } = null!;

    public string Pdoano { get; set; } = null!;

    public string Mesatu { get; set; } = null!;

    public string TpomonFnc { get; set; } = null!;

    public string? TpomonSgnMn { get; set; }

    public string? TpomonSgnMe { get; set; }

    public short CodctaNv3 { get; set; }

    public short CodctaNv4 { get; set; }

    public short CodctaNv5 { get; set; }

    public short CodctaNv6 { get; set; }

    public short CodctaNv7 { get; set; }

    public short CodctaNv8 { get; set; }

    public string? CodtdcPcp { get; set; }

    public string? CodtdcRtc { get; set; }

    public string? CodctaPcp { get; set; }

    public string? CodctaRtc { get; set; }

    public short Indcco { get; set; }

    public short Indmne { get; set; }

    public string? Indrtc { get; set; }

    public string? Indpcp { get; set; }

    public short CodccoNv3 { get; set; }

    public short CodccoNv5 { get; set; }

    public string TpogloRtc { get; set; } = null!;

    public string? GlodocrRtc { get; set; }

    public string? GlodocnRtc { get; set; }

    public string? CoddroIng { get; set; }

    public string? CoddroEgr { get; set; }

    public string? SernumeraRtc { get; set; }

    public string? NumerainiRtc { get; set; }

    public string? NumerafinRtc { get; set; }

    public string? AutosunatRtc { get; set; }

    public short NumeraDtr { get; set; }

    public string? CorrelainiRtc { get; set; }

    public short Ejerfran { get; set; }

    public string Indpedido { get; set; } = null!;

    public short Prodestino { get; set; }
}
