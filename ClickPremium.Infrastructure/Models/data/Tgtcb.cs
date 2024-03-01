using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de tipo de cambio
/// </summary>
public partial class Tgtcb
{
    public DateTime Fehtcb { get; set; }

    public double ImptcbCpr { get; set; }

    public double ImptcbVta { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
