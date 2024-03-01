using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de billetes y monedas
/// </summary>
public partial class Pldstmonedum
{
    public string Codmon { get; set; } = null!;

    public decimal Valordmo { get; set; }

    public string Desdmo { get; set; } = null!;

    public string Estadodmo { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
