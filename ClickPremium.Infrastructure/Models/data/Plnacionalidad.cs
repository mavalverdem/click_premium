using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Nacionalidad
/// </summary>
public partial class Plnacionalidad
{
    public string Codnac { get; set; } = null!;

    public string Desnac { get; set; } = null!;

    public string? Codpemi { get; set; }

    public string Estadonac { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
