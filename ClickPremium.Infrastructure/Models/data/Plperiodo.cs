using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de periodos de pago
/// </summary>
public partial class Plperiodo
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Despdo { get; set; } = null!;

    public string? Tpopdo { get; set; }

    public DateTime? Fechaini { get; set; }

    public DateTime? Fechafin { get; set; }

    public string? Anopdo { get; set; }

    public string? Mespdo { get; set; }

    public DateTime? Fechaproceso { get; set; }

    public DateTime? Fechapago { get; set; }

    public decimal Tipocambio { get; set; }

    public string Estadopdo { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Plasistencium> Plasistencia { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plcuentacte> PlcuentacteCodNavigations { get; set; } = new List<Plcuentacte>();

    public virtual ICollection<Plcuentacte> PlcuentacteCods { get; set; } = new List<Plcuentacte>();

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();

    public virtual ICollection<Plremuexce> Plremuexces { get; set; } = new List<Plremuexce>();
}
