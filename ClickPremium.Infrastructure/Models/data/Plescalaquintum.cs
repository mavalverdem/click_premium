using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Configuracion escala de rango quinta categoria
/// </summary>
public partial class Plescalaquintum
{
    public string Pdoanyo { get; set; } = null!;

    public short Orden { get; set; }

    public short Numerouit { get; set; }

    public decimal Factor { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
