using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de regimen pensionario - AFP
/// </summary>
public partial class Plentidadafp
{
    public string Codafp { get; set; } = null!;

    public string Desafp { get; set; } = null!;

    public decimal Factor1 { get; set; }

    public decimal Factor2 { get; set; }

    public decimal Factor3 { get; set; }

    public decimal Factor4 { get; set; }

    public string? Codbco { get; set; }

    public string? Ctacteafp { get; set; }

    public string? Desctacteafp { get; set; }

    public string? Ctactefondo { get; set; }

    public string? Desctactefondo { get; set; }

    public string Codsunat { get; set; } = null!;

    public string Estadoafp { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plctacenco> Plctacencos { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
