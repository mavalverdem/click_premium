using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Empresas que Destaca o Desplaza Personal
/// </summary>
public partial class Plempresasqde
{
    public string Codeqd { get; set; } = null!;

    public string Deseqd { get; set; } = null!;

    public string? Acteqd { get; set; }

    public DateTime? FechainiEqd { get; set; }

    public DateTime? FechafinEqd { get; set; }

    public string Estadoeqd { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
