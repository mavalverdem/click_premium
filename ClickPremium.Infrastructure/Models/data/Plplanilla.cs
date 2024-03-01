using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Formato de planilla ministerio
/// </summary>
public partial class Plplanilla
{
    public string Codcls { get; set; } = null!;

    public string Codpll { get; set; } = null!;

    public short Fila { get; set; }

    public short Columna { get; set; }

    public string Despll { get; set; } = null!;

    /// <summary>
    /// C:oncepto, D:ato
    /// </summary>
    public string? Tipo { get; set; }

    public string? Alias { get; set; }

    public string? Descripcion { get; set; }

    public short Posicion { get; set; }

    public short Longitud { get; set; }

    public string Subrayado { get; set; } = null!;

    public decimal Sizefont { get; set; }

    public string Sizepapel { get; set; } = null!;

    public string Posipapel { get; set; } = null!;

    public string Imprimecab { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Pldetaplanilla> Pldetaplanillas { get; set; } = new List<Pldetaplanilla>();
}
