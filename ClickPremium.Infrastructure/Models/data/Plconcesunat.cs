using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Conceptos Sunat
/// </summary>
public partial class Plconcesunat
{
    public string Codcon { get; set; } = null!;

    public string Descon { get; set; } = null!;

    public string Tipcon { get; set; } = null!;

    public string Segreg { get; set; } = null!;

    public string Segregcbssp { get; set; } = null!;

    public string Segagracui { get; set; } = null!;

    public string Sctr { get; set; } = null!;

    public string Ies { get; set; } = null!;

    public string Fdsa { get; set; } = null!;

    public string Senati { get; set; } = null!;

    public string Pensiones { get; set; } = null!;

    public string Spp { get; set; } = null!;

    public string Fcjmms { get; set; } = null!;

    public string Reptp { get; set; } = null!;

    public string Quinta { get; set; } = null!;

    public string Segregpen { get; set; } = null!;

    public string Csap { get; set; } = null!;

    public string Estadocon { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plconceplanilla> Plconceplanillas { get; set; } = new List<Plconceplanilla>();
}
