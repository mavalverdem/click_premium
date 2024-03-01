using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Permisos por usuario
/// </summary>
public partial class Sgpm
{
    public string Codusr { get; set; } = null!;

    public string Codemp { get; set; } = null!;

    public string Codsis { get; set; } = null!;

    public string Codmdl { get; set; } = null!;

    public short Indpms01 { get; set; }

    public short Indpms02 { get; set; }

    public short Indpms03 { get; set; }

    public short Indpms04 { get; set; }

    public short Indpms05 { get; set; }

    public short Indpms06 { get; set; }

    public short Indpms07 { get; set; }

    public short Indpms08 { get; set; }

    public short Indpms09 { get; set; }

    public short Indpms10 { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Tgemp CodempNavigation { get; set; } = null!;

    public virtual Sgusr CodusrNavigation { get; set; } = null!;
}
