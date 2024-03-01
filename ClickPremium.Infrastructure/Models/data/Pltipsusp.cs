using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo Suspension Laboral
/// </summary>
public partial class Pltipsusp
{
    public string Codtsu { get; set; } = null!;

    public string Destsu { get; set; } = null!;

    public string Estadotsu { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiAccs { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiEnferNavigations { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiFalleNavigations { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiNatalNavigations { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiPaterNavigations { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plasistencium> PlasistenciumCodmdiVacacNavigations { get; set; } = new List<Plasistencium>();
}
