using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Detalle de formatos de boleta de pago
/// </summary>
public partial class Pldetaboletum
{
    public string Codcls { get; set; } = null!;

    public string Codboleta { get; set; } = null!;

    public string Seccion { get; set; } = null!;

    public string Dato { get; set; } = null!;

    public string Tipodato { get; set; } = null!;

    public short Fila { get; set; }

    public short Columna { get; set; }

    public short Longitud { get; set; }

    public string Origen { get; set; } = null!;

    public decimal Sizefont { get; set; }

    public string Fontn { get; set; } = null!;

    public string Fonts { get; set; } = null!;

    public string Fontc { get; set; } = null!;

    public string? Desdato { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plboletapago Cod { get; set; } = null!;
}
