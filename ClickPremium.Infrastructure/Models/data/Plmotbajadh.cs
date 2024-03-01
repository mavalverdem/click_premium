using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Motivo de Baja de Derecho Habiente
/// </summary>
public partial class Plmotbajadh
{
    public string Codbdh { get; set; } = null!;

    public string Desbdh { get; set; } = null!;

    public string Estadobdh { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
