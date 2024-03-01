using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Vinculo Familiar
/// </summary>
public partial class Plvinfami
{
    public string Codvfa { get; set; } = null!;

    public string Desvfa { get; set; } = null!;

    public string Estadovfa { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
