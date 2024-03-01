using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de tipo de via
/// </summary>
public partial class Pltipovium
{
    public string Codvia { get; set; } = null!;

    public string Desvia { get; set; } = null!;

    public string? Abrevia { get; set; }

    public string Estadovia { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plcfgempresa> Plcfgempresas { get; set; } = new List<Plcfgempresa>();
}
