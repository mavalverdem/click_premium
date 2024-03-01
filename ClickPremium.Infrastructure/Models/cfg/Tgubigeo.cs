using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Maestro de ubicacion geografica
/// </summary>
public partial class Tgubigeo
{
    public string Codubg { get; set; } = null!;

    public string? Desubg { get; set; }

    public short? Nivelubg { get; set; }

    public string? Postalubg { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
