using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro sub periodos CTS
/// </summary>
public partial class Plctsperiodosub
{
    public string Codcls { get; set; } = null!;

    public string Pdocts { get; set; } = null!;

    public string Subcts { get; set; } = null!;

    public string Descrisub { get; set; } = null!;

    public string Pdoano { get; set; } = null!;

    public string Pdomes { get; set; } = null!;

    public short Numeroanos { get; set; }

    public short Numeromeses { get; set; }

    public short Numerodias { get; set; }

    public DateTime? Fechaini { get; set; }

    public DateTime? Fechafin { get; set; }

    public DateTime? Fechaven { get; set; }

    public DateTime? Fechacan { get; set; }

    public string Estadosub { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plctsmovimiento> Plctsmovimientos { get; set; } = new List<Plctsmovimiento>();

    public virtual Plctsperiodo Plctsperiodo { get; set; } = null!;
}
