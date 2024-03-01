using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Empresa que Desatacan o Desplazan Personal
/// </summary>
public partial class Plempresasqmde
{
    public string Codqmd { get; set; } = null!;

    public string Desqmd { get; set; } = null!;

    public string? Actqmd { get; set; }

    public DateTime? FechainiQmd { get; set; }

    public DateTime? FechafinQmd { get; set; }

    public string Estadoqmd { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
