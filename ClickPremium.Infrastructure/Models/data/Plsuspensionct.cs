using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Suspension de Cuarta Categoria
/// </summary>
public partial class Plsuspensionct
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string Numero { get; set; } = null!;

    public DateTime Fecha { get; set; }

    public string Ejercicio { get; set; } = null!;

    public string Medio { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
