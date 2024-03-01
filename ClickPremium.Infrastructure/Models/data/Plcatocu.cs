using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de categoría oupacional del trabajador
/// </summary>
public partial class Plcatocu
{
    public string Codcao { get; set; } = null!;

    public string Descao { get; set; } = null!;

    public string Estadocao { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plpersonal> Plpersonals { get; set; } = new List<Plpersonal>();
}
