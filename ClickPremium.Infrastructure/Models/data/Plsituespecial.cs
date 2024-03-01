using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Situación Especial
/// </summary>
public partial class Plsituespecial
{
    public string Codsie { get; set; } = null!;

    public string Dessie { get; set; } = null!;

    public string Estadosie { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
