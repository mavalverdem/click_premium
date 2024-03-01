using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de procesos de calculo
/// </summary>
public partial class Plproceso
{
    public string Codcls { get; set; } = null!;

    public string Codproce { get; set; } = null!;

    public string Desproce { get; set; } = null!;

    public string Estadoproce { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Plconceproceso> Plconceprocesos { get; set; } = new List<Plconceproceso>();
}
