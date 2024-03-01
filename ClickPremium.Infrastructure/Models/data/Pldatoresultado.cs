using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Datos de proceso de calculo
/// </summary>
public partial class Pldatoresultado
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string? Codcco { get; set; }

    public string? Codafp { get; set; }

    public string? Codeps { get; set; }

    public string Regpension { get; set; } = null!;

    public string? Naciextrapsn { get; set; }

    public DateTime? Fecingreso { get; set; }

    public string? Codubica { get; set; }

    public string? Codsec { get; set; }

    public string? Codcgo { get; set; }

    public string? Codcdt { get; set; }

    public DateTime? Fecestado { get; set; }

    public string Estadopsn { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plperiodo Cod { get; set; } = null!;

    public virtual Plpersonal CodNavigation { get; set; } = null!;

    public virtual Plentidadafp? CodafpNavigation { get; set; }

    public virtual Plconditrabajo? Codc { get; set; }

    public virtual Plcargo? CodcNavigation { get; set; }

    public virtual Cocco? CodccoNavigation { get; set; }

    public virtual Plentidadep? CodepsNavigation { get; set; }

    public virtual Plseccion? CodsecNavigation { get; set; }

    public virtual Plubicacion? CodubicaNavigation { get; set; }
}
