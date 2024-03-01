using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de secciones de personal
/// </summary>
public partial class Plseccion
{
    public string Codsec { get; set; } = null!;

    public string Dessec { get; set; } = null!;

    public string? Codintersec { get; set; }

    public string Estadosec { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plctacenco> Plctacencos { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctapv> Plctapvs { get; set; } = new List<Plctapv>();

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
