using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Control de Obligaciones sunat - cabecera
/// </summary>
public partial class Tgctrobli
{
    public string Pdotribu { get; set; } = null!;

    public DateTime? FecVence0 { get; set; }

    public DateTime? FecVence1 { get; set; }

    public DateTime? FecVence2 { get; set; }

    public DateTime? FecVence3 { get; set; }

    public DateTime? FecVence4 { get; set; }

    public DateTime? FecVence5 { get; set; }

    public DateTime? FecVence6 { get; set; }

    public DateTime? FecVence7 { get; set; }

    public DateTime? FecVence8 { get; set; }

    public DateTime? FecVence9 { get; set; }

    public DateTime? Buencontri { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Tgctroblidet> Tgctroblidets { get; set; } = new List<Tgctroblidet>();
}
