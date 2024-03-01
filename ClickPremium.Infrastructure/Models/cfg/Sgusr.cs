using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Maestro de usuarios
/// </summary>
public partial class Sgusr
{
    public string Codusr { get; set; } = null!;

    public string? Abvusr { get; set; }

    public string? Empusr { get; set; }

    public string Clausr { get; set; } = null!;

    public string? Nomusr { get; set; }

    public string? Anousr { get; set; }

    public string? Mesusr { get; set; }

    public string? Nvlusr { get; set; }

    public string Estusr { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Tgemp? EmpusrNavigation { get; set; }

    public virtual ICollection<Sgpm> Sgpms { get; set; } = new List<Sgpm>();
}
