using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Cuenta corriente de adelanto o prestamo
/// </summary>
public partial class Plcuentacte
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Numctacte { get; set; } = null!;

    public short Numcuota { get; set; }

    public string Tpoctacte { get; set; } = null!;

    public string? Codcpc { get; set; }

    public string? Codpdoprv { get; set; }

    public DateTime Fectacte { get; set; }

    public string Indchecar { get; set; } = null!;

    public string? Numchecar { get; set; }

    public string? Codbco { get; set; }

    public string Indgratifi { get; set; } = null!;

    public string Tpodscto { get; set; } = null!;

    public string Codmon { get; set; } = null!;

    public decimal CargoMn { get; set; }

    public decimal AbonoMn { get; set; }

    public decimal CargoMe { get; set; }

    public decimal AbonoMe { get; set; }

    public string Indprn { get; set; } = null!;

    public string? Codpdocan { get; set; }

    public string Estadoctacte { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plperiodo? Cod { get; set; }

    public virtual Plpersonal Cod1 { get; set; } = null!;

    public virtual Plperiodo? CodNavigation { get; set; }

    public virtual Plbanco? CodbcoNavigation { get; set; }

    public virtual Plconceplanilla? Codc { get; set; }
}
