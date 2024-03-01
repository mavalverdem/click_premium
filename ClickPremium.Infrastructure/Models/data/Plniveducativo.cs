using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Situación Educativa
/// </summary>
public partial class Plniveducativo
{
    public string Codniv { get; set; } = null!;

    public string Desniv { get; set; } = null!;

    public string Estadoniv { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
