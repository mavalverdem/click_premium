using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Remuneraciones o eventos excepcionales
/// </summary>
public partial class Plremuexce
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public string Codmon { get; set; } = null!;

    public decimal Imporemune { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plperiodo Cod { get; set; } = null!;

    public virtual Plpersonal CodNavigation { get; set; } = null!;

    public virtual Plconceplanilla Codc { get; set; } = null!;
}
