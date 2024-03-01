using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Tipo Documento Vínculo Familiar
/// </summary>
public partial class Pldocvinfami
{
    public string Coddvifa { get; set; } = null!;

    public string Desdvifa { get; set; } = null!;

    public string Estadotsu { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
