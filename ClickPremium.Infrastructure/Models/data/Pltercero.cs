using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Personal de Terceros
/// </summary>
public partial class Pltercero
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string? Ano { get; set; }

    public string? Mes { get; set; }

    public string? Ruc { get; set; }

    public string? Codest { get; set; }

    public string? Sctrs { get; set; }

    public string? Sctrp { get; set; }

    public decimal Tasa { get; set; }

    public decimal Importe { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
