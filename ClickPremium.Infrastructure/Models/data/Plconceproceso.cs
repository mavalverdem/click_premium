using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

public partial class Plconceproceso
{
    public string Codcls { get; set; } = null!;

    public string Codproce { get; set; } = null!;

    public string Codcpc { get; set; } = null!;

    public int Secuencia { get; set; }

    public string? Formulafun { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plproceso Cod { get; set; } = null!;

    public virtual Plconceplanilla Codc { get; set; } = null!;
}
