using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Detalle Formato de planilla ministerio
/// </summary>
public partial class Pldetaplanilla
{
    public string Codcls { get; set; } = null!;

    public string Codpll { get; set; } = null!;

    public short Fila { get; set; }

    public short Columna { get; set; }

    public string Codcpc { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plconceplanilla Codc { get; set; } = null!;

    public virtual Plplanilla Plplanilla { get; set; } = null!;
}
