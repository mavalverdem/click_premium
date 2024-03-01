using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo de Contrato de Trabajo
/// </summary>
public partial class Pltipcontrato
{
    public string Codtco { get; set; } = null!;

    public string Destco { get; set; } = null!;

    public string Estadotco { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plcontrato> Plcontratos { get; set; } = new List<Plcontrato>();
}
