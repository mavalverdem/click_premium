using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Remuneración Default Trabajador
/// </summary>
public partial class Plremudefa
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public string Codmon { get; set; } = null!;

    public decimal Imporemune { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;

    public virtual Plconceplanilla Codc { get; set; } = null!;
}
