using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo de Modalidad Formativa Laboral
/// </summary>
public partial class Plmodforma
{
    public string Codmfo { get; set; } = null!;

    public string Desmfo { get; set; } = null!;

    public string Estadomfo { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
