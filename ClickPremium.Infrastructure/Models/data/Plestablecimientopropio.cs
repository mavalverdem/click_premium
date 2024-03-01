using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Establecimientos Propios
/// </summary>
public partial class Plestablecimientopropio
{
    public string Codepr { get; set; } = null!;

    public string Tipepr { get; set; } = null!;

    public string? Cdgepr { get; set; }

    public string? Desepr { get; set; }

    public string? Indepr { get; set; }

    public decimal Tasepr { get; set; }

    public string Estadoepr { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
