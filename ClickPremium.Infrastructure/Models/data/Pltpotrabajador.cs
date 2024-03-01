using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de tipo de trabajador
/// </summary>
public partial class Pltpotrabajador
{
    public string Codtpt { get; set; } = null!;

    public string Destpt { get; set; } = null!;

    public string Estadotpt { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
