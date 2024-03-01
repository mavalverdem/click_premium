using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Comprobantes de Cuarta categoria
/// </summary>
public partial class Plcomprobantect
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string Tipo { get; set; } = null!;

    public string Serie { get; set; } = null!;

    public string Numero { get; set; } = null!;

    public decimal Monto { get; set; }

    public DateTime Fecemision { get; set; }

    public DateTime Fecpago { get; set; }

    public string Retencion { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;
}
