using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de cuentas provisión
/// </summary>
public partial class Plctapv
{
    public string Codcls { get; set; } = null!;

    public string Codcco { get; set; } = null!;

    public string Codsec { get; set; } = null!;

    public short Orden { get; set; }

    public string? CodctavacDebmn { get; set; }

    public string? CodctavacHabmn { get; set; }

    public string? CodctavacDebme { get; set; }

    public string? CodctavacHabme { get; set; }

    public string? CodctavexDebmn { get; set; }

    public string? CodctavexHabmn { get; set; }

    public string? CodctavexDebme { get; set; }

    public string? CodctavexHabme { get; set; }

    public string? CodctagraDebmn { get; set; }

    public string? CodctagraHabmn { get; set; }

    public string? CodctagraDebme { get; set; }

    public string? CodctagraHabme { get; set; }

    public string? CodctagexDebmn { get; set; }

    public string? CodctagexHabmn { get; set; }

    public string? CodctagexDebme { get; set; }

    public string? CodctagexHabme { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Cocco CodccoNavigation { get; set; } = null!;

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual Coctum? CodctagexDebmeNavigation { get; set; }

    public virtual Coctum? CodctagexDebmnNavigation { get; set; }

    public virtual Coctum? CodctagexHabmeNavigation { get; set; }

    public virtual Coctum? CodctagexHabmnNavigation { get; set; }

    public virtual Coctum? CodctagraDebmeNavigation { get; set; }

    public virtual Coctum? CodctagraDebmnNavigation { get; set; }

    public virtual Coctum? CodctagraHabmeNavigation { get; set; }

    public virtual Coctum? CodctagraHabmnNavigation { get; set; }

    public virtual Coctum? CodctavacDebmeNavigation { get; set; }

    public virtual Coctum? CodctavacDebmnNavigation { get; set; }

    public virtual Coctum? CodctavacHabmeNavigation { get; set; }

    public virtual Coctum? CodctavacHabmnNavigation { get; set; }

    public virtual Coctum? CodctavexDebmeNavigation { get; set; }

    public virtual Coctum? CodctavexDebmnNavigation { get; set; }

    public virtual Coctum? CodctavexHabmeNavigation { get; set; }

    public virtual Coctum? CodctavexHabmnNavigation { get; set; }

    public virtual Plseccion CodsecNavigation { get; set; } = null!;
}
