using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Maestro de tablas anexos sunat
/// </summary>
public partial class Tgsunat
{
    public string Codtabla { get; set; } = null!;

    public string Codsunat { get; set; } = null!;

    public string? Detsunat { get; set; }

    public string? Campo03 { get; set; }

    public string Estsunat { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
