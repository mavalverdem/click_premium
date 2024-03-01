using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Situacion Trabajador Pensionista
/// </summary>
public partial class Plsitrapen
{
    public string Codstp { get; set; } = null!;

    public string Desstp { get; set; } = null!;

    public string Estadostp { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
