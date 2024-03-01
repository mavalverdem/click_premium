using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Rango de impresiones
/// </summary>
public partial class Rangoimpresion
{
    public string Proceso { get; set; } = null!;

    public string Valor { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public string Fyhcre { get; set; } = null!;
}
