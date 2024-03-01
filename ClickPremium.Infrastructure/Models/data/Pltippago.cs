using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo de Pago
/// </summary>
public partial class Pltippago
{
    public string Codtip { get; set; } = null!;

    public string Destip { get; set; } = null!;

    public string Estadotip { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
