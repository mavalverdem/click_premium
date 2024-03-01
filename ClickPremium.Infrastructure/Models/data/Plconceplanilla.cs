using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Conceptos de calculo por clase planilla
/// </summary>
public partial class Plconceplanilla
{
    public string Codcls { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public string Clasecpc { get; set; } = null!;

    public string Defaultcpc { get; set; } = null!;

    public string Impbolecpc { get; set; } = null!;

    public string? Formulafun { get; set; }

    public string? Imagenfun { get; set; }

    public string? Codsunat { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual Plconcepto CodcpcNavigation { get; set; } = null!;

    public virtual Plconcesunat? CodsunatNavigation { get; set; }

    public virtual ICollection<Plcartabanco> Plcartabancos { get; set; } = new List<Plcartabanco>();

    public virtual ICollection<Plconceproceso> Plconceprocesos { get; set; } = new List<Plconceproceso>();

    public virtual ICollection<Plctacenco> Plctacencos { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctsresultado> Plctsresultados { get; set; } = new List<Plctsresultado>();

    public virtual ICollection<Plcuentacte> Plcuentactes { get; set; } = new List<Plcuentacte>();

    public virtual ICollection<Pldetaplanilla> Pldetaplanillas { get; set; } = new List<Pldetaplanilla>();

    public virtual ICollection<Plremudefa> Plremudefas { get; set; } = new List<Plremudefa>();

    public virtual ICollection<Plremuexce> Plremuexces { get; set; } = new List<Plremuexce>();
}
