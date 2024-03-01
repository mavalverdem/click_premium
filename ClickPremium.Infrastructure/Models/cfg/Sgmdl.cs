using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Modulos u opciones del sistema
/// </summary>
public partial class Sgmdl
{
    public string Codsis { get; set; } = null!;

    public string Opcion { get; set; } = null!;

    public string Orden { get; set; } = null!;

    public string Codmdl { get; set; } = null!;

    public string Detmdl { get; set; } = null!;

    public string? Detmdlx { get; set; }

    public string? Nommdl { get; set; }

    public string Estmdl { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
