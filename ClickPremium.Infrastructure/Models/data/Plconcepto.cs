using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de conceptos de calculo
/// </summary>
public partial class Plconcepto
{
    public string Codcpc { get; set; } = null!;

    public string Descpc { get; set; } = null!;

    public string? Aliascpc { get; set; }

    public string Tipocpc { get; set; } = null!;

    public string? Obs { get; set; }

    public string Estadocpc { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plconceplanilla> Plconceplanillas { get; set; } = new List<Plconceplanilla>();
}
