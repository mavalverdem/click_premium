using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Resultado de Procesos de CTS
/// </summary>
public partial class Plctsresultado
{
    public string Codcls { get; set; } = null!;

    public string Pdocts { get; set; } = null!;

    public string Subcts { get; set; } = null!;

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

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plconceplanilla Codc { get; set; } = null!;

    public virtual Coctum? CodctaDebmeNavigation { get; set; }

    public virtual Coctum? CodctaDebmnNavigation { get; set; }

    public virtual Coctum? CodctaHabmeNavigation { get; set; }

    public virtual Coctum? CodctaHabmnNavigation { get; set; }

    public virtual Plctsmovimiento Plctsmovimiento { get; set; } = null!;
}
