using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Motivo Fin Contrato o baja TRegistro
/// </summary>
public partial class Plmotfin
{
    public string Codmof { get; set; } = null!;

    public string Desmof { get; set; } = null!;

    public string Estadomof { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
