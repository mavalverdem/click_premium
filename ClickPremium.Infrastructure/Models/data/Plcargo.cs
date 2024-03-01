using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de cargo de personal
/// </summary>
public partial class Plcargo
{
    public string Codcls { get; set; } = null!;

    public string Codcgo { get; set; } = null!;

    public string Descgo { get; set; } = null!;

    public string Estadocgo { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();

    public virtual ICollection<Plexpelaboral> Plexpelaborals { get; set; } = new List<Plexpelaboral>();

    public virtual ICollection<Plpersonal> Plpersonals { get; set; } = new List<Plpersonal>();
}
