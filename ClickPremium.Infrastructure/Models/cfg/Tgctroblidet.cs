using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Control de Obligaciones sunat - detalle
/// </summary>
public partial class Tgctroblidet
{
    public string Codemp { get; set; } = null!;

    public string Pdotribu { get; set; } = null!;

    public string? Coddeclar { get; set; }

    public string? Nroconsta { get; set; }

    public DateTime? Fpresenta { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Tgemp CodempNavigation { get; set; } = null!;

    public virtual Tgctrobli PdotribuNavigation { get; set; } = null!;
}
