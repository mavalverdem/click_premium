using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de Plan de Cuentas
/// </summary>
public partial class Coctum
{
    public string Codcta { get; set; } = null!;

    public string Detcta { get; set; } = null!;

    public string? Detctax { get; set; }

    public short Tpocta { get; set; }

    public short Natcta { get; set; }

    public string? Tposdo { get; set; }

    public string? Tpoanl { get; set; }

    public string? CodctaDstDeb { get; set; }

    public string? CodctaDstHabb { get; set; }

    public string? CodccoDstDeb { get; set; }

    public string? CodccoDstHab { get; set; }

    public string Tpomon { get; set; } = null!;

    public string Tpotcb { get; set; } = null!;

    public string? Tpoajd { get; set; }

    public string? CodctaAjdDeb { get; set; }

    public string? CodctaAjdHab { get; set; }

    public string? CodccoAjdDeb { get; set; }

    public string? CodccoAjdHab { get; set; }

    public short Indajd { get; set; }

    public string? CodctaCrrDeu { get; set; }

    public string? CodctaCrrAcr { get; set; }

    public string? CodccoDef { get; set; }

    public short Indcco { get; set; }

    public short Inddoc { get; set; }

    public short Indnoe { get; set; }

    public short Indpsp { get; set; }

    public short Indfjo { get; set; }

    public string? Codbco { get; set; }

    public string Estcta { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual ICollection<Plctacenco> PlctacencoCodctaDebmeNavigations { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctacenco> PlctacencoCodctaDebmnNavigations { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctacenco> PlctacencoCodctaHabmeNavigations { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctacenco> PlctacencoCodctaHabmnNavigations { get; set; } = new List<Plctacenco>();

    public virtual ICollection<Plctapv> PlctapvCodctagexDebmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagexDebmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagexHabmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagexHabmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagraDebmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagraDebmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagraHabmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctagraHabmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavacDebmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavacDebmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavacHabmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavacHabmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavexDebmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavexDebmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavexHabmeNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctapv> PlctapvCodctavexHabmnNavigations { get; set; } = new List<Plctapv>();

    public virtual ICollection<Plctsresultado> PlctsresultadoCodctaDebmeNavigations { get; set; } = new List<Plctsresultado>();

    public virtual ICollection<Plctsresultado> PlctsresultadoCodctaDebmnNavigations { get; set; } = new List<Plctsresultado>();

    public virtual ICollection<Plctsresultado> PlctsresultadoCodctaHabmeNavigations { get; set; } = new List<Plctsresultado>();

    public virtual ICollection<Plctsresultado> PlctsresultadoCodctaHabmnNavigations { get; set; } = new List<Plctsresultado>();
}
