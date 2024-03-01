using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Cabecera de generador de reporte
/// </summary>
public partial class Plgenreporte
{
    public string Codcls { get; set; } = null!;

    public string Codrpt { get; set; } = null!;

    public string Desrpt { get; set; } = null!;

    public string? Formarpt { get; set; }

    public string? Titulorpt { get; set; }

    public string? Pierpt { get; set; }

    public string Interlinea { get; set; } = null!;

    public short Anchorpt { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Pldetareporte> Pldetareportes { get; set; } = new List<Pldetareporte>();
}
