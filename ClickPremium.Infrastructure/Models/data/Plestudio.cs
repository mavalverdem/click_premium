using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de estudio realizado
/// </summary>
public partial class Plestudio
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string? Institucion { get; set; }

    public string? Grado { get; set; }

    public DateTime? Fechaini { get; set; }

    public DateTime? Fechafin { get; set; }

    public string? Observacion { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;
}
