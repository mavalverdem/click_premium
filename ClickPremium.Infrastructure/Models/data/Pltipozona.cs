using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de tipo de zona
/// </summary>
public partial class Pltipozona
{
    public string Codzona { get; set; } = null!;

    public string Deszona { get; set; } = null!;

    public string? Abrezona { get; set; }

    public string Estadozona { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plcfgempresa> Plcfgempresas { get; set; } = new List<Plcfgempresa>();
}
