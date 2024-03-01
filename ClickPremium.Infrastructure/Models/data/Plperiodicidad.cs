using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro periodicidad de pago
/// </summary>
public partial class Plperiodicidad
{
    public string Codprd { get; set; } = null!;

    public string Desprd { get; set; } = null!;

    public string Estadoprd { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
