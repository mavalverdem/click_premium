using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de periodos de calculo CTS
/// </summary>
public partial class Plctsperiodo
{
    public string Codcls { get; set; } = null!;

    public string Pdocts { get; set; } = null!;

    public string Descricts { get; set; } = null!;

    public string Pdoano { get; set; } = null!;

    public string Estadocts { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Plctsperiodosub> Plctsperiodosubs { get; set; } = new List<Plctsperiodosub>();
}
