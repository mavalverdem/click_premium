using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de centro de costos
/// </summary>
public partial class Cocco
{
    public string Codcco { get; set; } = null!;

    public string Detcco { get; set; } = null!;

    public string? Detccox { get; set; }

    public string Estcco { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plcfgcencosto? Plcfgcencosto { get; set; }

    public virtual ICollection<Plctacenco> Plctacencos { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctapv> Plctapvs { get; set; } = new List<Plctapv>();

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
