using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Pais Emisor Documento
/// </summary>
public partial class Plpaisemidocum
{
    public string Codpemi { get; set; } = null!;

    public string Despemi { get; set; } = null!;

    public string Estadopemi { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
