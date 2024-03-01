using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro de personal
/// </summary>
public partial class Plpersonal
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public string? Apepaterno { get; set; }

    public string? Apematerno { get; set; }

    public string? Nombres { get; set; }

    public DateTime Fecnacimiento { get; set; }

    public string? Ubigeonac { get; set; }

    public string? Nacionalidad { get; set; }

    public string? Codpemi { get; set; }

    public string Naciextrapsn { get; set; } = null!;

    public string? Sexopsn { get; set; }

    public string? Codniv { get; set; }

    public string? Refedirec { get; set; }

    public string? Codvia { get; set; }

    public string? Nomviadirec { get; set; }

    public string? Numerdirec { get; set; }

    public string? Intedirec { get; set; }

    public string? Codzona { get; set; }

    public string? Nomzondirec { get; set; }

    public string? Ubigeodir { get; set; }

    public string? Estcivilpsn { get; set; }

    public short Numhijo { get; set; }

    public short Numdepen { get; set; }

    public string? Coddci { get; set; }

    public string? Numdociden { get; set; }

    public string? Numdocmil { get; set; }

    public string? Codldn { get; set; }

    public string? Telefono { get; set; }

    public string? Celular { get; set; }

    public string Dctojudicial { get; set; } = null!;

    public decimal Pordsctojudi { get; set; }

    public DateTime Fecingreso { get; set; }

    public string Reingreso { get; set; } = null!;

    public string? Codtpt { get; set; }

    public string? Codcgo { get; set; }

    public string Cgoconfianza { get; set; } = null!;

    public string? Codpfs { get; set; }

    public decimal Jornadalaboral { get; set; }

    public string? Codcco { get; set; }

    public string? Codcdt { get; set; }

    public string? Codafp { get; set; }

    public string? Numeroafp { get; set; }

    public string Afpmixta { get; set; } = null!;

    public string Pagodolar { get; set; } = null!;

    public string? Periodicidad { get; set; }

    public string? Tippago { get; set; }

    public string? Codbcopago { get; set; }

    public string? Cuentapago { get; set; }

    public string? Interbankpago { get; set; }

    public string? Codbnkpago { get; set; }

    public string? Ctsdeposito { get; set; }

    public string? Ctsdolar { get; set; }

    public string? Codbcocts { get; set; }

    public string? Cuentacts { get; set; }

    public string? Interbankcts { get; set; }

    public string? Codbnkcts { get; set; }

    public string? Cuentaibankcts { get; set; }

    public string? Codeps { get; set; }

    public string? Regpension { get; set; }

    public DateTime? Fecingregpen { get; set; }

    public string? Essvida { get; set; }

    public string? Cobsctr { get; set; }

    public string? Afilsindical { get; set; }

    public string? Remintegralgrati { get; set; }

    public string? Remintegralvaca { get; set; }

    public string? Remintegralcts { get; set; }

    public string? Remimprecisa { get; set; }

    public string? Remuneta { get; set; }

    public string? Netocpc { get; set; }

    public string? Variacpc { get; set; }

    public decimal? Imporemuneto { get; set; }

    public DateTime? Fecbaja { get; set; }

    public string? Nroessalud { get; set; }

    public string? Codubica { get; set; }

    public string? Codsec { get; set; }

    public string? Coddeudor { get; set; }

    public string? Codacredor { get; set; }

    public DateTime? Fecestado { get; set; }

    public byte[]? Fotopsn { get; set; }

    public string? Correoelect { get; set; }

    public string ChkSctrp { get; set; } = null!;

    public string ChkRl { get; set; } = null!;

    public string ChkDis { get; set; } = null!;

    public string ChkMax { get; set; } = null!;

    public string ChkReg { get; set; } = null!;

    public string ChkNoc { get; set; } = null!;

    public string ChkQui { get; set; } = null!;

    public string ChkOiq { get; set; } = null!;

    public string? Siteps { get; set; }

    public string Segmedico { get; set; } = null!;

    public string Resfamiliar { get; set; } = null!;

    public string Forprofesional { get; set; } = null!;

    public string? Finperiodo { get; set; }

    public string? Modformativa { get; set; }

    public string ChkPe { get; set; } = null!;

    public string? Cmbcatocupacional { get; set; }

    public string? Cmbtributacion { get; set; }

    public string Chk27252 { get; set; } = null!;

    public string Estadopsn { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Plcatocu? CmbcatocupacionalNavigation { get; set; }

    public virtual Plconven? CmbtributacionNavigation { get; set; }

    public virtual Plbanco? CodbcoctsNavigation { get; set; }

    public virtual Plbanco? CodbcopagoNavigation { get; set; }

    public virtual Plbanco? CodbnkctsNavigation { get; set; }

    public virtual Plbanco? CodbnkpagoNavigation { get; set; }

    public virtual Plcargo? Codc { get; set; }

    public virtual Plclasplan CodclsNavigation { get; set; } = null!;

    public virtual ICollection<Plasistencium> Plasistencia { get; set; } = new List<Plasistencium>();

    public virtual ICollection<Plcartabanco> Plcartabancos { get; set; } = new List<Plcartabanco>();

    public virtual ICollection<Plcomprobantect> Plcomprobantects { get; set; } = new List<Plcomprobantect>();

    public virtual ICollection<Plcontrato> Plcontratos { get; set; } = new List<Plcontrato>();

    public virtual ICollection<Plctsmovimiento> Plctsmovimientos { get; set; } = new List<Plctsmovimiento>();

    public virtual ICollection<Plcuentacte> Plcuentactes { get; set; } = new List<Plcuentacte>();

    public virtual ICollection<Pldatoresultado> Pldatoresultados { get; set; } = new List<Pldatoresultado>();

    public virtual ICollection<Plestudio> Plestudios { get; set; } = new List<Plestudio>();

    public virtual ICollection<Plexpelaboral> Plexpelaborals { get; set; } = new List<Plexpelaboral>();

    public virtual ICollection<Plremudefa> Plremudefas { get; set; } = new List<Plremudefa>();

    public virtual ICollection<Plremuexce> Plremuexces { get; set; } = new List<Plremuexce>();
}
