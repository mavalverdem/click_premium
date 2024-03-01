using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models;

/// <summary>
/// Maestro de empresas
/// </summary>
public partial class Tgemp
{
    public string Codemp { get; set; } = null!;

    public string? Razemp { get; set; }

    public string? Rucemp { get; set; }

    public string? Direccion { get; set; }

    public string? Localidademp { get; set; }

    public string? Actividademp { get; set; }

    public string? Repapepaterno { get; set; }

    public string? Repapematerno { get; set; }

    public string? Repnombre { get; set; }

    public string? Repdocumento { get; set; }

    public string? Conapepaterno { get; set; }

    public string? Conapematerno { get; set; }

    public string? Connombre { get; set; }

    public string? Condocumento { get; set; }

    public string Buencontri { get; set; } = null!;

    public string? Indret { get; set; }

    public string? Indper { get; set; }

    public string? CodctaRet { get; set; }

    public string? CodctaPer { get; set; }

    public string? DetMn { get; set; }

    public string? SmbMn { get; set; }

    public string? DetMe { get; set; }

    public string? SmbMe { get; set; }

    public string Siscon { get; set; } = null!;

    public string Sispla { get; set; } = null!;

    public string Sisban { get; set; } = null!;

    public string? Actividad { get; set; }

    public string Sincroniza { get; set; } = null!;

    public string? Empcon { get; set; }

    public string? Nombredbemp { get; set; }

    public string Estemp { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Sgpm> Sgpms { get; set; } = new List<Sgpm>();

    public virtual ICollection<Sgusr> Sgusrs { get; set; } = new List<Sgusr>();

    public virtual ICollection<Tgctroblidet> Tgctroblidets { get; set; } = new List<Tgctroblidet>();
}
