using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Movimiento CTS por personal
/// </summary>
public partial class Plctsmovimiento
{
    public string Codcls { get; set; } = null!;

    public string Pdocts { get; set; } = null!;

    public string Subcts { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Pdoano { get; set; } = null!;

    public string Pdomes { get; set; } = null!;

    public short Numeroanos { get; set; }

    public short Numeromeses { get; set; }

    public short Numerodias { get; set; }

    public DateTime? Fechaini { get; set; }

    public DateTime? Fechafin { get; set; }

    public DateTime? Fechaven { get; set; }

    public DateTime? Fechacan { get; set; }

    public decimal Porinteres { get; set; }

    public decimal Tipocambio { get; set; }

    public string? Nrodeposito { get; set; }

    public string Estadomov { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;

    public virtual Plctsperiodosub Plctsperiodosub { get; set; } = null!;

    public virtual ICollection<Plctsresultado> Plctsresultados { get; set; } = new List<Plctsresultado>();
}
