using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo de Establecimiento
/// </summary>
public partial class Plestablecimiento
{
    public string Codest { get; set; } = null!;

    public string Desest { get; set; } = null!;

    public string Estadoest { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
