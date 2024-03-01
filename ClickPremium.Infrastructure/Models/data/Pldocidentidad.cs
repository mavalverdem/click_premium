using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de documento de identidad
/// </summary>
public partial class Pldocidentidad
{
    public string Coddci { get; set; } = null!;

    public string Desdci { get; set; } = null!;

    public string? Sigladci { get; set; }

    public string Codsunat { get; set; } = null!;

    public string Estadodci { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plcfgempresa> PlcfgempresaGercoddciNavigations { get; set; } = new List<Plcfgempresa>();

    public virtual ICollection<Plcfgempresa> PlcfgempresaRepcoddciNavigations { get; set; } = new List<Plcfgempresa>();
}
