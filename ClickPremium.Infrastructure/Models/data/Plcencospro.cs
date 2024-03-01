using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Distribución Centro de Costos
/// </summary>
public partial class Plcencospro
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Codcco { get; set; } = null!;

    public decimal Porcentaje { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
