using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Resultado proceso calculo planilla
/// </summary>
public partial class Plresultado
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Codproce { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public int Secuencia { get; set; }

    public string? Codmon { get; set; }

    public decimal ImporteMn { get; set; }

    public decimal ImporteMe { get; set; }

    public string? CodctaDebmn { get; set; }

    public string? CodctaHabmn { get; set; }

    public string? CodctaDebme { get; set; }

    public string? CodctaHabme { get; set; }

    public string? Pdoano { get; set; }

    public string? Pdomes { get; set; }

    public string Tipocpc { get; set; } = null!;

    public string Impbolecpc { get; set; } = null!;

    public string? CodprocePdo { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
