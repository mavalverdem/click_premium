using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de contrato de trabajo
/// </summary>
public partial class Plcontrato
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Numdocumen { get; set; } = null!;

    public string Anyo { get; set; } = null!;

    public string Mes { get; set; } = null!;

    public string Dia { get; set; } = null!;

    public DateTime? Fechaini { get; set; }

    public DateTime? Fechafin { get; set; }

    public string? Observacion { get; set; }

    public string? Archivo { get; set; }

    public string? Tipcon { get; set; }

    public string Estadocon { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;

    public virtual Pltipcontrato? TipconNavigation { get; set; }
}
