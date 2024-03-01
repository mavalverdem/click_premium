using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Condicion de Trabajo
/// </summary>
public partial class Plconditrabajo
{
    public string Codcls { get; set; } = null!;

    public string Codcdt { get; set; } = null!;

    public string Descdt { get; set; } = null!;

    public string Estadocdt { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
