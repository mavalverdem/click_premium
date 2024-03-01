using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Formato de boleta de pago
/// </summary>
public partial class Plboletapago
{
    public string Codcls { get; set; } = null!;

    public string Codboleta { get; set; } = null!;

    public string Desboleta { get; set; } = null!;

    public string Orientacion { get; set; } = null!;

    public string Calidad { get; set; } = null!;

    public decimal Papelancho { get; set; }

    public decimal Papelalto { get; set; }

    public string? Font { get; set; }

    public string Copia { get; set; } = null!;

    public short Lininicopia { get; set; }

    public string Estadobol { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Pldetaboletum> Pldetaboleta { get; set; } = new List<Pldetaboletum>();
}
