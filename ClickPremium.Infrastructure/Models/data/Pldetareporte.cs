using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Detalle formato de generador de reporte
/// </summary>
public partial class Pldetareporte
{
    public string Codcls { get; set; } = null!;

    public string Codrpt { get; set; } = null!;

    public short Orden { get; set; }

    public string Descripcion { get; set; } = null!;

    /// <summary>
    /// A:cumulador, C:oncepto, D:ato
    /// </summary>
    public string? Tipo { get; set; }

    public string? Alias { get; set; }

    public short Nivel { get; set; }

    public string Signo { get; set; } = null!;

    public string Impreso { get; set; } = null!;

    public short Longitud { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plgenreporte Cod { get; set; } = null!;
}
