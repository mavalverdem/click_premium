using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo Comprobante Servicio
/// </summary>
public partial class Pltipcom
{
    public string Codtic { get; set; } = null!;

    public string Destic { get; set; } = null!;

    public string Estadotic { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
