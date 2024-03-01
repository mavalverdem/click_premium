using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Convenios evitar doble tributación
/// </summary>
public partial class Plconven
{
    public string Codctr { get; set; } = null!;

    public string Desctr { get; set; } = null!;

    public string Estadoctr { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plpersonal> Plpersonals { get; set; } = new List<Plpersonal>();
}
