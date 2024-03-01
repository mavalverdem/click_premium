using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de clase de planilla
/// </summary>
public partial class Plclasplan
{
    public string Codcls { get; set; } = null!;

    public string Descls { get; set; } = null!;

    public string? Clave { get; set; }

    public decimal Horadiaria { get; set; }

    public string Fmtboleta { get; set; } = null!;

    public string Fmtrecibo { get; set; } = null!;

    public string Tipo { get; set; } = null!;

    public string Estadocls { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plboletapago> Plboletapagos { get; set; } = new List<Plboletapago>();

    public virtual ICollection<Plcargo> Plcargos { get; set; } = new List<Plcargo>();

    public virtual ICollection<Plconceplanilla> Plconceplanillas { get; set; } = new List<Plconceplanilla>();

    public virtual ICollection<Plconditrabajo> Plconditrabajos { get; set; } = new List<Plconditrabajo>();

    public virtual ICollection<Plctapv> Plctapvs { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctsperiodo> Plctsperiodos { get; set; } = new List<Plctsperiodo>();

    public virtual ICollection<Plgenreporte> Plgenreportes { get; set; } = new List<Plgenreporte>();

    public virtual ICollection<Plperiodo> Plperiodos { get; set; } = new List<Plperiodo>();

    public virtual ICollection<Plpersonal> Plpersonals { get; set; } = new List<Plpersonal>();

    public virtual ICollection<Plplanilla> Plplanillas { get; set; } = new List<Plplanilla>();

    public virtual ICollection<Plproceso> Plprocesos { get; set; } = new List<Plproceso>();
}
