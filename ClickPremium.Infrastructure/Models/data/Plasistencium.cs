using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Control de asistencia del personal
/// </summary>
public partial class Plasistencium
{
    public string Codcls { get; set; } = null!;

    public string Codpdo { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Diatrabajo { get; set; }

    public short Diamediotm { get; set; }

    public short Diaparcial { get; set; }

    public short Dialaboral { get; set; }

    public decimal Horanormal { get; set; }

    public decimal Horamediotm { get; set; }

    public decimal Horaparcial { get; set; }

    public decimal Horatipo1 { get; set; }

    public decimal Horatipo2 { get; set; }

    public decimal Horatipo3 { get; set; }

    public decimal Horatipo4 { get; set; }

    public short Diafalta { get; set; }

    public decimal Tardanza { get; set; }

    public short Diaprepostnatal { get; set; }

    public string? CodmdiNatal { get; set; }

    public DateTime? FechainiNatal { get; set; }

    public DateTime? FechafinNatal { get; set; }

    public string? NumecittNatal { get; set; }

    public short Accidente { get; set; }

    public string? CodmdiAccid { get; set; }

    public DateTime? FechainiAccid { get; set; }

    public DateTime? FechafinAccid { get; set; }

    public short Diavacaciones { get; set; }

    public string? CodmdiVacac { get; set; }

    public short Enfermedad { get; set; }

    public string? CodmdiEnfer { get; set; }

    public DateTime? FechainiEnfer { get; set; }

    public DateTime? FechafinEnfer { get; set; }

    public string? NumecittEnfer { get; set; }

    public short Licencia { get; set; }

    public string? CodmdiLicen { get; set; }

    public DateTime? FechainiLicen { get; set; }

    public DateTime? FechafinLicen { get; set; }

    public short Diaferiado { get; set; }

    public short Diatradesemanal { get; set; }

    public short Diasuspension { get; set; }

    public short Dialibre { get; set; }

    public decimal Permisos { get; set; }

    public DateTime? Fechainivacacion { get; set; }

    public DateTime? Fechafinvacacion { get; set; }

    public string? Pdovaca1 { get; set; }

    public DateTime? Fechainivaca1 { get; set; }

    public DateTime? Fechafinvaca1 { get; set; }

    public string? Pdovaca2 { get; set; }

    public DateTime? Fechainivaca2 { get; set; }

    public DateTime? Fechafinvaca2 { get; set; }

    public short Dialiquidacion { get; set; }

    public short Liquidavacacion { get; set; }

    public short Diagratificacion { get; set; }

    public DateTime? Fechacese { get; set; }

    public DateTime? Fechainiliqvaca { get; set; }

    public DateTime? Fechafinliqvaca { get; set; }

    public string? Observacion { get; set; }

    public short Liqnocalifica { get; set; }

    public short Tercerturno { get; set; }

    public decimal Opcional { get; set; }

    public short Diavacaventa { get; set; }

    public string? Pdovaca3 { get; set; }

    public DateTime? Fechainivaca3 { get; set; }

    public DateTime? Fechafinvaca3 { get; set; }

    public string Indvacadelanta { get; set; } = null!;

    public decimal? Diavacavencida { get; set; }

    public string? Pdovaca4 { get; set; }

    public DateTime? Fechainivaca4 { get; set; }

    public DateTime? Fechafinvaca4 { get; set; }

    public string? Pdovaca5 { get; set; }

    public DateTime? Fechainivaca5 { get; set; }

    public DateTime? Fechafinvaca5 { get; set; }

    public short Diapaternidad { get; set; }

    public string? CodmdiPater { get; set; }

    public DateTime? FechainiPater { get; set; }

    public DateTime? FechafinPater { get; set; }

    public short Diafallecefam { get; set; }

    public string? CodmdiFalle { get; set; }

    public DateTime? FechainiFalle { get; set; }

    public DateTime? FechafinFalle { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plperiodo Cod { get; set; } = null!;

    public virtual Plpersonal CodNavigation { get; set; } = null!;

    public virtual Pltipsusp? CodmdiAcc { get; set; }

    public virtual Pltipsusp? CodmdiEnferNavigation { get; set; }

    public virtual Pltipsusp? CodmdiFalleNavigation { get; set; }

    public virtual Pltipsusp? CodmdiNatalNavigation { get; set; }

    public virtual Pltipsusp? CodmdiPaterNavigation { get; set; }

    public virtual Pltipsusp? CodmdiVacacNavigation { get; set; }
}
