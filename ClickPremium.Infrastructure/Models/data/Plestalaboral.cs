using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Establecimiento Labora Trabajador
/// </summary>
public partial class Plestalaboral
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string Ano { get; set; } = null!;

    public string Mes { get; set; } = null!;

    public string Ruc { get; set; } = null!;

    public string? Codest { get; set; }

    public decimal? Tasa { get; set; }

    public string? Usrcre { get; set; }

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
