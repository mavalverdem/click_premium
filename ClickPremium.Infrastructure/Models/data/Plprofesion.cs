using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de profesiones u ocupaciones
/// </summary>
public partial class Plprofesion
{
    public string Codpfs { get; set; } = null!;

    public string Despfs { get; set; } = null!;

    public string Cateobrpfs { get; set; } = null!;

    public string Estadopfs { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
