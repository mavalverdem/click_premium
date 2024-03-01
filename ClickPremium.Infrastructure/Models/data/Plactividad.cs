using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maesto Tipo Actividad Empresarial SUNAT
/// </summary>
public partial class Plactividad
{
    public string Codact { get; set; } = null!;

    public string Desact { get; set; } = null!;

    public string Estadoact { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
