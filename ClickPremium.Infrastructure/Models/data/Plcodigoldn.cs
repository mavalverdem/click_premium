using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de código larga distancia nacional
/// </summary>
public partial class Plcodigoldn
{
    public string Codldn { get; set; } = null!;

    public string Desldn { get; set; } = null!;

    public string Estadoldn { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
