using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de entidades bancarias
/// </summary>
public partial class Plbanco
{
    public string Codbco { get; set; } = null!;

    public string Desbco { get; set; } = null!;

    public string? Cuentamn { get; set; }

    public string? Cuentame { get; set; }

    public string? Codentidad { get; set; }

    public string? Formato { get; set; }

    public decimal ImpolimiteMn { get; set; }

    public decimal ImpolimiteMe { get; set; }

    public string Estadobco { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plcartabanco> Plcartabancos { get; set; } = new List<Plcartabanco>();

    public virtual ICollection<Plcuentacte> Plcuentactes { get; set; } = new List<Plcuentacte>();

    public virtual ICollection<Plpersonal> PlpersonalCodbcoctsNavigations { get; set; } = new List<Plpersonal>();

    public virtual ICollection<Plpersonal> PlpersonalCodbcopagoNavigations { get; set; } = new List<Plpersonal>();

    public virtual ICollection<Plpersonal> PlpersonalCodbnkctsNavigations { get; set; } = new List<Plpersonal>();

    public virtual ICollection<Plpersonal> PlpersonalCodbnkpagoNavigations { get; set; } = new List<Plpersonal>();
}
