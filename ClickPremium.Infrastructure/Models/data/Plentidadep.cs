using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de empresas prestadora de servicios
/// </summary>
public partial class Plentidadep
{
    public string Codeps { get; set; } = null!;

    public string Deseps { get; set; } = null!;

    public string? Ruceps { get; set; }

    public decimal Factoreps { get; set; }

    public string Codsunat { get; set; } = null!;

    public string Estadoeps { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();
}
