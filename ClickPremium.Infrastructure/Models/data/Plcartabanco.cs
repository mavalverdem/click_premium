using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Historico de transferencias de bancos
/// </summary>
public partial class Plcartabanco
{
    public string Codcls { get; set; } = null!;

    public string Codbco { get; set; } = null!;

    public string Nrocarta { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string Desmotivo { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public DateTime Fechaproce { get; set; }

    public string Codmon { get; set; } = null!;

    public decimal ImporteMn { get; set; }

    public decimal ImporteMe { get; set; }

    public decimal Porinteres { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plpersonal Cod { get; set; } = null!;

    public virtual Plbanco CodbcoNavigation { get; set; } = null!;

    public virtual Plconceplanilla Codc { get; set; } = null!;
}
