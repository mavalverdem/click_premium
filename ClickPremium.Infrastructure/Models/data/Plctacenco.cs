using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de cuentas por centro de costo
/// </summary>
public partial class Plctacenco
{
    public string Codcls { get; set; } = null!;

    public string Codcco { get; set; } = null!;

    public string Codsec { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public short Orden { get; set; }

    public string? Codafp { get; set; }

    public string? CodctaDebmn { get; set; }

    public string? CodctaHabmn { get; set; }

    public string? CodctaDebme { get; set; }

    public string? CodctaHabme { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plentidadafp? CodafpNavigation { get; set; }

    public virtual Plconceplanilla Codc { get; set; } = null!;

    public virtual Cocco CodccoNavigation { get; set; } = null!;

    public virtual Coctum? CodctaDebmeNavigation { get; set; }

    public virtual Coctum? CodctaDebmnNavigation { get; set; }

    public virtual Coctum? CodctaHabmeNavigation { get; set; }

    public virtual Coctum? CodctaHabmnNavigation { get; set; }

    public virtual Plseccion CodsecNavigation { get; set; } = null!;
}
