using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de ubicacion o localidad
/// </summary>
public partial class Plubicacion
{
    public string Codubica { get; set; } = null!;

    public string Desubica { get; set; } = null!;

    public string? Codinterubica { get; set; }

    public string Estadoubica { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
