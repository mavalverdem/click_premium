using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace ClickPremium.Infrastructure.Models.data;

public partial class ClickpremSysmaplaContext : DbContext
{
    public ClickpremSysmaplaContext()
    {
    }

    public ClickpremSysmaplaContext(DbContextOptions<ClickpremSysmaplaContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Cocco> Coccos { get; set; }

    public virtual DbSet<Cocfg> Cocfgs { get; set; }

    public virtual DbSet<Coctum> Cocta { get; set; }

    public virtual DbSet<Plactividad> Plactividads { get; set; }

    public virtual DbSet<Plasistencium> Plasistencia { get; set; }

    public virtual DbSet<Plbanco> Plbancos { get; set; }

    public virtual DbSet<Plboletapago> Plboletapagos { get; set; }

    public virtual DbSet<Plcargo> Plcargos { get; set; }

    public virtual DbSet<Plcartabanco> Plcartabancos { get; set; }

    public virtual DbSet<Plcatocu> Plcatocus { get; set; }

    public virtual DbSet<Plcencospro> Plcencospros { get; set; }

    public virtual DbSet<Plcfgcencosto> Plcfgcencostos { get; set; }

    public virtual DbSet<Plcfgempresa> Plcfgempresas { get; set; }

    public virtual DbSet<Plclasplan> Plclasplans { get; set; }

    public virtual DbSet<Plcodigoldn> Plcodigoldns { get; set; }

    public virtual DbSet<Plcomprobantect> Plcomprobantects { get; set; }

    public virtual DbSet<Plconceplanilla> Plconceplanillas { get; set; }

    public virtual DbSet<Plconceproceso> Plconceprocesos { get; set; }

    public virtual DbSet<Plconcepto> Plconceptos { get; set; }

    public virtual DbSet<Plconcesunat> Plconcesunats { get; set; }

    public virtual DbSet<Plconditrabajo> Plconditrabajos { get; set; }

    public virtual DbSet<Plcontrato> Plcontratos { get; set; }

    public virtual DbSet<Plconven> Plconvens { get; set; }

    public virtual DbSet<Plctacenco> Plctacencos { get; set; }

    public virtual DbSet<Plctapv> Plctapvs { get; set; }

    public virtual DbSet<Plctsmovimiento> Plctsmovimientos { get; set; }

    public virtual DbSet<Plctsperiodo> Plctsperiodos { get; set; }

    public virtual DbSet<Plctsperiodosub> Plctsperiodosubs { get; set; }

    public virtual DbSet<Plctsresultado> Plctsresultados { get; set; }

    public virtual DbSet<Plcuentacte> Plcuentactes { get; set; }

    public virtual DbSet<Pldatoresultado> Pldatoresultados { get; set; }

    public virtual DbSet<Pldetaboletum> Pldetaboleta { get; set; }

    public virtual DbSet<Pldetaplanilla> Pldetaplanillas { get; set; }

    public virtual DbSet<Pldetareporte> Pldetareportes { get; set; }

    public virtual DbSet<Pldocidentidad> Pldocidentidads { get; set; }

    public virtual DbSet<Pldocvinfami> Pldocvinfamis { get; set; }

    public virtual DbSet<Pldstmonedum> Pldstmoneda { get; set; }

    public virtual DbSet<Plempleadore> Plempleadores { get; set; }

    public virtual DbSet<Plempresaseqde> Plempresaseqdes { get; set; }

    public virtual DbSet<Plempresasqde> Plempresasqdes { get; set; }

    public virtual DbSet<Plempresasqmde> Plempresasqmdes { get; set; }

    public virtual DbSet<Plentidadafp> Plentidadafps { get; set; }

    public virtual DbSet<Plentidadep> Plentidadeps { get; set; }

    public virtual DbSet<Plescalaquintum> Plescalaquinta { get; set; }

    public virtual DbSet<Plestablecimiento> Plestablecimientos { get; set; }

    public virtual DbSet<Plestablecimientopropio> Plestablecimientopropios { get; set; }

    public virtual DbSet<Plestalaboral> Plestalaborals { get; set; }

    public virtual DbSet<Plestudio> Plestudios { get; set; }

    public virtual DbSet<Plexpelaboral> Plexpelaborals { get; set; }

    public virtual DbSet<Plfamiliare> Plfamiliares { get; set; }

    public virtual DbSet<Plgenreporte> Plgenreportes { get; set; }

    public virtual DbSet<Plmodforma> Plmodformas { get; set; }

    public virtual DbSet<Plmotbajadh> Plmotbajadhs { get; set; }

    public virtual DbSet<Plmotfin> Plmotfins { get; set; }

    public virtual DbSet<Plnacionalidad> Plnacionalidads { get; set; }

    public virtual DbSet<Plniveducativo> Plniveducativos { get; set; }

    public virtual DbSet<Plpaisemidocum> Plpaisemidocums { get; set; }

    public virtual DbSet<Plperiodicidad> Plperiodicidads { get; set; }

    public virtual DbSet<Plperiodo> Plperiodos { get; set; }

    public virtual DbSet<Plpersonal> Plpersonals { get; set; }

    public virtual DbSet<Plplanilla> Plplanillas { get; set; }

    public virtual DbSet<Plproceso> Plprocesos { get; set; }

    public virtual DbSet<Plprofesion> Plprofesions { get; set; }

    public virtual DbSet<Plremudefa> Plremudefas { get; set; }

    public virtual DbSet<Plremuexce> Plremuexces { get; set; }

    public virtual DbSet<Plresultado> Plresultados { get; set; }

    public virtual DbSet<Plseccion> Plseccions { get; set; }

    public virtual DbSet<Plsitrapen> Plsitrapens { get; set; }

    public virtual DbSet<Plsituespecial> Plsituespecials { get; set; }

    public virtual DbSet<Plsuspensionct> Plsuspensioncts { get; set; }

    public virtual DbSet<Pltercero> Plterceros { get; set; }

    public virtual DbSet<Pltipcom> Pltipcoms { get; set; }

    public virtual DbSet<Pltipcontrato> Pltipcontratos { get; set; }

    public virtual DbSet<Pltipovium> Pltipovia { get; set; }

    public virtual DbSet<Pltipozona> Pltipozonas { get; set; }

    public virtual DbSet<Pltippago> Pltippagos { get; set; }

    public virtual DbSet<Pltipsusp> Pltipsusps { get; set; }

    public virtual DbSet<Pltpotrabajador> Pltpotrabajadors { get; set; }

    public virtual DbSet<Plubicacion> Plubicacions { get; set; }

    public virtual DbSet<Plvarfunc> Plvarfuncs { get; set; }

    public virtual DbSet<Plvinfami> Plvinfamis { get; set; }

    public virtual DbSet<Rangoimpresion> Rangoimpresions { get; set; }

    public virtual DbSet<Tgtcb> Tgtcbs { get; set; }

//     protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
// #warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
//         => optionsBuilder.UseSqlServer("Server=localhost;Database=CLICKPREM_sysmapla;User Id=sa;Password=P@ssw0rd@DB1;TrustServerCertificate=true");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Cocco>(entity =>
        {
            entity.HasKey(e => e.Codcco).HasName("PK_cocco_codcco");

            entity.ToTable("cocco", tb => tb.HasComment("Maestro de centro de costos"));

            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Detcco)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("detcco");
            entity.Property(e => e.Detccox)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("detccox");
            entity.Property(e => e.Estcco)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estcco");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Cocfg>(entity =>
        {
            entity.HasKey(e => new { e.Codemp, e.Pdoano }).HasName("PK_cocfg_empanyo");

            entity.ToTable("cocfg", tb => tb.HasComment("Parametros de configuracion"));

            entity.Property(e => e.Codemp)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codemp");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdoano");
            entity.Property(e => e.AutosunatRtc)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("autosunat_rtc");
            entity.Property(e => e.CodccoNv3).HasColumnName("codcco_nv3");
            entity.Property(e => e.CodccoNv5).HasColumnName("codcco_nv5");
            entity.Property(e => e.CodctaNv3).HasColumnName("codcta_nv3");
            entity.Property(e => e.CodctaNv4).HasColumnName("codcta_nv4");
            entity.Property(e => e.CodctaNv5).HasColumnName("codcta_nv5");
            entity.Property(e => e.CodctaNv6).HasColumnName("codcta_nv6");
            entity.Property(e => e.CodctaNv7).HasColumnName("codcta_nv7");
            entity.Property(e => e.CodctaNv8).HasColumnName("codcta_nv8");
            entity.Property(e => e.CodctaPcp)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codcta_pcp");
            entity.Property(e => e.CodctaRtc)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codcta_rtc");
            entity.Property(e => e.CoddroEgr)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("coddro_egr");
            entity.Property(e => e.CoddroIng)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("coddro_ing");
            entity.Property(e => e.CodtdcPcp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtdc_pcp");
            entity.Property(e => e.CodtdcRtc)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtdc_rtc");
            entity.Property(e => e.CorrelainiRtc)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("correlaini_rtc");
            entity.Property(e => e.Ejerfran).HasColumnName("ejerfran");
            entity.Property(e => e.GlodocnRtc)
                .HasMaxLength(250)
                .IsUnicode(false)
                .HasColumnName("glodocn_rtc");
            entity.Property(e => e.GlodocrRtc)
                .HasMaxLength(250)
                .IsUnicode(false)
                .HasColumnName("glodocr_rtc");
            entity.Property(e => e.Indcco).HasColumnName("indcco");
            entity.Property(e => e.Indmne).HasColumnName("indmne");
            entity.Property(e => e.Indpcp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indpcp");
            entity.Property(e => e.Indpedido)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indpedido");
            entity.Property(e => e.Indrtc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indrtc");
            entity.Property(e => e.Mesatu)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mesatu");
            entity.Property(e => e.NumeraDtr).HasColumnName("numera_dtr");
            entity.Property(e => e.NumerafinRtc)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("numerafin_rtc");
            entity.Property(e => e.NumerainiRtc)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("numeraini_rtc");
            entity.Property(e => e.Prodestino).HasColumnName("prodestino");
            entity.Property(e => e.SernumeraRtc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("sernumera_rtc");
            entity.Property(e => e.TpogloRtc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpoglo_rtc");
            entity.Property(e => e.TpomonFnc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpomon_fnc");
            entity.Property(e => e.TpomonSgnMe)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("tpomon_sgn_me");
            entity.Property(e => e.TpomonSgnMn)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("tpomon_sgn_mn");
        });

        modelBuilder.Entity<Coctum>(entity =>
        {
            entity.HasKey(e => e.Codcta).HasName("PK_cocta_codcta");

            entity.ToTable("cocta", tb => tb.HasComment("Maestro de Plan de Cuentas"));

            entity.Property(e => e.Codcta)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta");
            entity.Property(e => e.Codbco)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbco");
            entity.Property(e => e.CodccoAjdDeb)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco_ajd_deb");
            entity.Property(e => e.CodccoAjdHab)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco_ajd_hab");
            entity.Property(e => e.CodccoDef)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco_def");
            entity.Property(e => e.CodccoDstDeb)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco_dst_deb");
            entity.Property(e => e.CodccoDstHab)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco_dst_hab");
            entity.Property(e => e.CodctaAjdDeb)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_ajd_deb");
            entity.Property(e => e.CodctaAjdHab)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_ajd_hab");
            entity.Property(e => e.CodctaCrrAcr)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_crr_acr");
            entity.Property(e => e.CodctaCrrDeu)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_crr_deu");
            entity.Property(e => e.CodctaDstDeb)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_dst_deb");
            entity.Property(e => e.CodctaDstHabb)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_dst_habb");
            entity.Property(e => e.Detcta)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("detcta");
            entity.Property(e => e.Detctax)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("detctax");
            entity.Property(e => e.Estcta)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estcta");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Indajd).HasColumnName("indajd");
            entity.Property(e => e.Indcco).HasColumnName("indcco");
            entity.Property(e => e.Inddoc).HasColumnName("inddoc");
            entity.Property(e => e.Indfjo).HasColumnName("indfjo");
            entity.Property(e => e.Indnoe).HasColumnName("indnoe");
            entity.Property(e => e.Indpsp).HasColumnName("indpsp");
            entity.Property(e => e.Natcta).HasColumnName("natcta");
            entity.Property(e => e.Tpoajd)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpoajd");
            entity.Property(e => e.Tpoanl)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpoanl");
            entity.Property(e => e.Tpocta).HasColumnName("tpocta");
            entity.Property(e => e.Tpomon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpomon");
            entity.Property(e => e.Tposdo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tposdo");
            entity.Property(e => e.Tpotcb)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpotcb");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plactividad>(entity =>
        {
            entity.HasKey(e => e.Codact).HasName("PK_plactividad_codact");

            entity.ToTable("plactividad", tb => tb.HasComment("Maesto Tipo Actividad Empresarial SUNAT"));

            entity.Property(e => e.Codact)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codact");
            entity.Property(e => e.Desact)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desact");
            entity.Property(e => e.Estadoact)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoact");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plasistencium>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo, e.Codpsn }).HasName("PK_plasistencia_clspdopsn");

            entity.ToTable("plasistencia", tb => tb.HasComment("Control de asistencia del personal"));

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "IX_plasistencia_clspsn");

            entity.HasIndex(e => e.CodmdiAccid, "IX_plasistencia_codmdi_accid");

            entity.HasIndex(e => e.CodmdiEnfer, "IX_plasistencia_codmdi_enfer");

            entity.HasIndex(e => e.CodmdiFalle, "IX_plasistencia_codmdi_falle");

            entity.HasIndex(e => e.CodmdiNatal, "IX_plasistencia_codmdi_natal");

            entity.HasIndex(e => e.CodmdiPater, "IX_plasistencia_codmdi_pater");

            entity.HasIndex(e => e.CodmdiVacac, "IX_plasistencia_codmdi_vacac");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Accidente).HasColumnName("accidente");
            entity.Property(e => e.CodmdiAccid)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_accid");
            entity.Property(e => e.CodmdiEnfer)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_enfer");
            entity.Property(e => e.CodmdiFalle)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_falle");
            entity.Property(e => e.CodmdiLicen)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_licen");
            entity.Property(e => e.CodmdiNatal)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_natal");
            entity.Property(e => e.CodmdiPater)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_pater");
            entity.Property(e => e.CodmdiVacac)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmdi_vacac");
            entity.Property(e => e.Diafallecefam).HasColumnName("diafallecefam");
            entity.Property(e => e.Diafalta).HasColumnName("diafalta");
            entity.Property(e => e.Diaferiado).HasColumnName("diaferiado");
            entity.Property(e => e.Diagratificacion).HasColumnName("diagratificacion");
            entity.Property(e => e.Dialaboral).HasColumnName("dialaboral");
            entity.Property(e => e.Dialibre).HasColumnName("dialibre");
            entity.Property(e => e.Dialiquidacion).HasColumnName("dialiquidacion");
            entity.Property(e => e.Diamediotm).HasColumnName("diamediotm");
            entity.Property(e => e.Diaparcial).HasColumnName("diaparcial");
            entity.Property(e => e.Diapaternidad).HasColumnName("diapaternidad");
            entity.Property(e => e.Diaprepostnatal).HasColumnName("diaprepostnatal");
            entity.Property(e => e.Diasuspension).HasColumnName("diasuspension");
            entity.Property(e => e.Diatrabajo).HasColumnName("diatrabajo");
            entity.Property(e => e.Diatradesemanal).HasColumnName("diatradesemanal");
            entity.Property(e => e.Diavacaciones).HasColumnName("diavacaciones");
            entity.Property(e => e.Diavacavencida)
                .HasColumnType("decimal(6, 3)")
                .HasColumnName("diavacavencida");
            entity.Property(e => e.Diavacaventa).HasColumnName("diavacaventa");
            entity.Property(e => e.Enfermedad).HasColumnName("enfermedad");
            entity.Property(e => e.Fechacese)
                .HasColumnType("date")
                .HasColumnName("fechacese");
            entity.Property(e => e.FechafinAccid)
                .HasColumnType("date")
                .HasColumnName("fechafin_accid");
            entity.Property(e => e.FechafinEnfer)
                .HasColumnType("date")
                .HasColumnName("fechafin_enfer");
            entity.Property(e => e.FechafinFalle)
                .HasColumnType("date")
                .HasColumnName("fechafin_falle");
            entity.Property(e => e.FechafinLicen)
                .HasColumnType("date")
                .HasColumnName("fechafin_licen");
            entity.Property(e => e.FechafinNatal)
                .HasColumnType("date")
                .HasColumnName("fechafin_natal");
            entity.Property(e => e.FechafinPater)
                .HasColumnType("date")
                .HasColumnName("fechafin_pater");
            entity.Property(e => e.Fechafinliqvaca)
                .HasColumnType("date")
                .HasColumnName("fechafinliqvaca");
            entity.Property(e => e.Fechafinvaca1)
                .HasColumnType("date")
                .HasColumnName("fechafinvaca1");
            entity.Property(e => e.Fechafinvaca2)
                .HasColumnType("date")
                .HasColumnName("fechafinvaca2");
            entity.Property(e => e.Fechafinvaca3)
                .HasColumnType("date")
                .HasColumnName("fechafinvaca3");
            entity.Property(e => e.Fechafinvaca4)
                .HasColumnType("date")
                .HasColumnName("fechafinvaca4");
            entity.Property(e => e.Fechafinvaca5)
                .HasColumnType("date")
                .HasColumnName("fechafinvaca5");
            entity.Property(e => e.Fechafinvacacion)
                .HasColumnType("date")
                .HasColumnName("fechafinvacacion");
            entity.Property(e => e.FechainiAccid)
                .HasColumnType("date")
                .HasColumnName("fechaini_accid");
            entity.Property(e => e.FechainiEnfer)
                .HasColumnType("date")
                .HasColumnName("fechaini_enfer");
            entity.Property(e => e.FechainiFalle)
                .HasColumnType("date")
                .HasColumnName("fechaini_falle");
            entity.Property(e => e.FechainiLicen)
                .HasColumnType("date")
                .HasColumnName("fechaini_licen");
            entity.Property(e => e.FechainiNatal)
                .HasColumnType("date")
                .HasColumnName("fechaini_natal");
            entity.Property(e => e.FechainiPater)
                .HasColumnType("date")
                .HasColumnName("fechaini_pater");
            entity.Property(e => e.Fechainiliqvaca)
                .HasColumnType("date")
                .HasColumnName("fechainiliqvaca");
            entity.Property(e => e.Fechainivaca1)
                .HasColumnType("date")
                .HasColumnName("fechainivaca1");
            entity.Property(e => e.Fechainivaca2)
                .HasColumnType("date")
                .HasColumnName("fechainivaca2");
            entity.Property(e => e.Fechainivaca3)
                .HasColumnType("date")
                .HasColumnName("fechainivaca3");
            entity.Property(e => e.Fechainivaca4)
                .HasColumnType("date")
                .HasColumnName("fechainivaca4");
            entity.Property(e => e.Fechainivaca5)
                .HasColumnType("date")
                .HasColumnName("fechainivaca5");
            entity.Property(e => e.Fechainivacacion)
                .HasColumnType("date")
                .HasColumnName("fechainivacacion");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Horamediotm)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horamediotm");
            entity.Property(e => e.Horanormal)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horanormal");
            entity.Property(e => e.Horaparcial)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horaparcial");
            entity.Property(e => e.Horatipo1)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horatipo1");
            entity.Property(e => e.Horatipo2)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horatipo2");
            entity.Property(e => e.Horatipo3)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horatipo3");
            entity.Property(e => e.Horatipo4)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horatipo4");
            entity.Property(e => e.Indvacadelanta)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indvacadelanta");
            entity.Property(e => e.Licencia).HasColumnName("licencia");
            entity.Property(e => e.Liqnocalifica).HasColumnName("liqnocalifica");
            entity.Property(e => e.Liquidavacacion).HasColumnName("liquidavacacion");
            entity.Property(e => e.NumecittEnfer)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("numecitt_enfer");
            entity.Property(e => e.NumecittNatal)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("numecitt_natal");
            entity.Property(e => e.Observacion)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("observacion");
            entity.Property(e => e.Opcional)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("opcional");
            entity.Property(e => e.Pdovaca1)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("pdovaca1");
            entity.Property(e => e.Pdovaca2)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("pdovaca2");
            entity.Property(e => e.Pdovaca3)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("pdovaca3");
            entity.Property(e => e.Pdovaca4)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("pdovaca4");
            entity.Property(e => e.Pdovaca5)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("pdovaca5");
            entity.Property(e => e.Permisos)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("permisos");
            entity.Property(e => e.Tardanza)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("tardanza");
            entity.Property(e => e.Tercerturno).HasColumnName("tercerturno");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodmdiAcc).WithMany(p => p.PlasistenciumCodmdiAccs).HasForeignKey(d => d.CodmdiAccid);

            entity.HasOne(d => d.CodmdiEnferNavigation).WithMany(p => p.PlasistenciumCodmdiEnferNavigations).HasForeignKey(d => d.CodmdiEnfer);

            entity.HasOne(d => d.CodmdiFalleNavigation).WithMany(p => p.PlasistenciumCodmdiFalleNavigations).HasForeignKey(d => d.CodmdiFalle);

            entity.HasOne(d => d.CodmdiNatalNavigation).WithMany(p => p.PlasistenciumCodmdiNatalNavigations).HasForeignKey(d => d.CodmdiNatal);

            entity.HasOne(d => d.CodmdiPaterNavigation).WithMany(p => p.PlasistenciumCodmdiPaterNavigations).HasForeignKey(d => d.CodmdiPater);

            entity.HasOne(d => d.CodmdiVacacNavigation).WithMany(p => p.PlasistenciumCodmdiVacacNavigations).HasForeignKey(d => d.CodmdiVacac);

            entity.HasOne(d => d.Cod).WithMany(p => p.Plasistencia)
                .HasForeignKey(d => new { d.Codcls, d.Codpdo })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plasistencia_plperiodo_clspdo");

            entity.HasOne(d => d.CodNavigation).WithMany(p => p.Plasistencia)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plasistencia_plpersonal_clspsn");
        });

        modelBuilder.Entity<Plbanco>(entity =>
        {
            entity.HasKey(e => e.Codbco).HasName("PK_plbanco_codbco");

            entity.ToTable("plbanco", tb => tb.HasComment("Maestro de entidades bancarias"));

            entity.Property(e => e.Codbco)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbco");
            entity.Property(e => e.Codentidad)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("codentidad");
            entity.Property(e => e.Cuentame)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cuentame");
            entity.Property(e => e.Cuentamn)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cuentamn");
            entity.Property(e => e.Desbco)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desbco");
            entity.Property(e => e.Estadobco)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadobco");
            entity.Property(e => e.Formato)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("formato");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.ImpolimiteMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("impolimite_me");
            entity.Property(e => e.ImpolimiteMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("impolimite_mn");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plboletapago>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codboleta }).HasName("PK_plboletapago_clsboleta");

            entity.ToTable("plboletapago", tb => tb.HasComment("Formato de boleta de pago"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codboleta)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codboleta");
            entity.Property(e => e.Calidad)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("calidad");
            entity.Property(e => e.Copia)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("copia");
            entity.Property(e => e.Desboleta)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desboleta");
            entity.Property(e => e.Estadobol)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadobol");
            entity.Property(e => e.Font)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("font");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Lininicopia).HasColumnName("lininicopia");
            entity.Property(e => e.Orientacion)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("orientacion");
            entity.Property(e => e.Papelalto)
                .HasColumnType("decimal(5, 2)")
                .HasColumnName("papelalto");
            entity.Property(e => e.Papelancho)
                .HasColumnType("decimal(5, 2)")
                .HasColumnName("papelancho");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plboletapagos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plcargo>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codcgo }).HasName("PK_plcargo_clscgo");

            entity.ToTable("plcargo", tb => tb.HasComment("Maestro de cargo de personal"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codcgo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcgo");
            entity.Property(e => e.Descgo)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descgo");
            entity.Property(e => e.Estadocgo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocgo");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plcargos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plcartabanco>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codbco, e.Nrocarta, e.Codcpc, e.Codpsn }).HasName("PK__plcartab__A779EDC0B9ECD19C");

            entity.ToTable("plcartabanco", tb => tb.HasComment("Historico de transferencias de bancos"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "IX_plcartabanco_clscpc");

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "IX_plcartabanco_clspsn");

            entity.HasIndex(e => e.Codbco, "IX_plcartabanco_codbco");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codbco)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbco");
            entity.Property(e => e.Nrocarta)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("nrocarta");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Desmotivo)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("desmotivo");
            entity.Property(e => e.Fechaproce)
                .HasColumnType("date")
                .HasColumnName("fechaproce");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.ImporteMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_me");
            entity.Property(e => e.ImporteMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_mn");
            entity.Property(e => e.Porinteres)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("porinteres");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodbcoNavigation).WithMany(p => p.Plcartabancos)
                .HasForeignKey(d => d.Codbco)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.Codc).WithMany(p => p.Plcartabancos)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plcartabanco_plconceplanilla_clscpc");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plcartabancos)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plcartabanco_plpersonal_clspsn");
        });

        modelBuilder.Entity<Plcatocu>(entity =>
        {
            entity.HasKey(e => e.Codcao).HasName("PK_plcatocu_codcao");

            entity.ToTable("plcatocu", tb => tb.HasComment("Maestro de categoría oupacional del trabajador"));

            entity.Property(e => e.Codcao)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcao");
            entity.Property(e => e.Descao)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descao");
            entity.Property(e => e.Estadocao)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocao");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plcencospro>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo, e.Codpsn, e.Codcco }).HasName("PK__plcencos__6A3ED20BBFA4C3E6");

            entity.ToTable("plcencospro", tb => tb.HasComment("Distribución Centro de Costos"));

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "IX_plcencospro_clspsn");

            entity.HasIndex(e => e.Codcco, "IX_plcencospro_codcco");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codcco)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Porcentaje)
                .HasColumnType("decimal(7, 2)")
                .HasColumnName("porcentaje");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plcfgcencosto>(entity =>
        {
            entity.HasKey(e => e.Codcco).HasName("PK_plcfgcencosto_codcco");

            entity.ToTable("plcfgcencosto", tb => tb.HasComment("Maestro Informacion del Negocio centro costo"));

            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Clientenego)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("clientenego");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Lineanegocio)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("lineanegocio");
            entity.Property(e => e.Segmentonego)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segmentonego");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodccoNavigation).WithOne(p => p.Plcfgcencosto)
                .HasForeignKey<Plcfgcencosto>(d => d.Codcco)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plcfgempresa>(entity =>
        {
            entity.HasKey(e => e.Pdoano).HasName("PK_plcfgempresa_pdoano");

            entity.ToTable("plcfgempresa", tb => tb.HasComment("Configuracion de parametros de empresa"));

            entity.HasIndex(e => e.Codvia, "IX_plcfgempresa_codvia");

            entity.HasIndex(e => e.Codzona, "IX_plcfgempresa_codzona");

            entity.HasIndex(e => e.Gercoddci, "IX_plcfgempresa_gercoddci");

            entity.HasIndex(e => e.Repcoddci, "IX_plcfgempresa_repcoddci");

            entity.HasIndex(e => e.Ubigeodir, "IX_plcfgempresa_ubigeodir");

            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Codcpc5ta)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc5ta");
            entity.Property(e => e.Codcpc5taIng)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc5ta_Ing");
            entity.Property(e => e.Codcpcrem)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpcrem");
            entity.Property(e => e.Codtbldividir)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtbldividir");
            entity.Property(e => e.Codtblpendiente)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtblpendiente");
            entity.Property(e => e.Codtblretener)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtblretener");
            entity.Property(e => e.Codtbluit)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtbluit");
            entity.Property(e => e.Codvia)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codvia");
            entity.Property(e => e.Codzona)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codzona");
            entity.Property(e => e.ContratoDoc)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("contrato_doc");
            entity.Property(e => e.ContratoDot)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("contrato_dot");
            entity.Property(e => e.CorreoEnvio)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("correo_envio");
            entity.Property(e => e.Direccionvia)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("direccionvia");
            entity.Property(e => e.Direccionzona)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("direccionzona");
            entity.Property(e => e.Dirimpbol)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("dirimpbol");
            entity.Property(e => e.Email)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("email");
            entity.Property(e => e.Firma).HasColumnName("firma");
            entity.Property(e => e.Firmanexo).HasColumnName("firmanexo");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Gerapematerno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("gerapematerno");
            entity.Property(e => e.Gerapepaterno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("gerapepaterno");
            entity.Property(e => e.Gercargo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("gercargo");
            entity.Property(e => e.Gercoddci)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("gercoddci");
            entity.Property(e => e.Gernombres)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("gernombres");
            entity.Property(e => e.Gernumdocu)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("gernumdocu");
            entity.Property(e => e.Girocomercial)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("girocomercial");
            entity.Property(e => e.Gratiliqxdias)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("gratiliqxdias");
            entity.Property(e => e.Gratipendiente)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("gratipendiente");
            entity.Property(e => e.Gratixasis)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("gratixasis");
            entity.Property(e => e.LiqprnLogoemp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("liqprn_logoemp");
            entity.Property(e => e.LiqprnRazonemp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("liqprn_razonemp");
            entity.Property(e => e.Logo).HasColumnName("logo");
            entity.Property(e => e.Nivelcencosto).HasColumnName("nivelcencosto");
            entity.Property(e => e.Numerodir)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("numerodir");
            entity.Property(e => e.PasswordEnvio)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("password_envio");
            entity.Property(e => e.Porcepartici)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("porcepartici");
            entity.Property(e => e.Psnapematerno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("psnapematerno");
            entity.Property(e => e.Psnapepaterno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("psnapepaterno");
            entity.Property(e => e.Psnnombres)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("psnnombres");
            entity.Property(e => e.Psntelefono)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("psntelefono");
            entity.Property(e => e.PuertoEnvio).HasColumnName("puerto_envio");
            entity.Property(e => e.Regpatronal)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("regpatronal");
            entity.Property(e => e.Remanterior)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remanterior");
            entity.Property(e => e.Rembasica)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("rembasica");
            entity.Property(e => e.Remempextra)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remempextra");
            entity.Property(e => e.Remempordin)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remempordin");
            entity.Property(e => e.Remganada)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remganada");
            entity.Property(e => e.Rempendiente)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("rempendiente");
            entity.Property(e => e.Rempromedio)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("rempromedio");
            entity.Property(e => e.Remxutiejer1)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remxutiejer1");
            entity.Property(e => e.Remxutiejer2)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remxutiejer2");
            entity.Property(e => e.Remxutiejer3)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remxutiejer3");
            entity.Property(e => e.Remxutiejer4)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("remxutiejer4");
            entity.Property(e => e.RentaxejerMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("rentaxejer_me");
            entity.Property(e => e.RentaxejerMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("rentaxejer_mn");
            entity.Property(e => e.Repapematerno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repapematerno");
            entity.Property(e => e.Repapepaterno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repapepaterno");
            entity.Property(e => e.Repcargo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("repcargo");
            entity.Property(e => e.Repcoddci)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("repcoddci");
            entity.Property(e => e.Repimpbol)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("repimpbol");
            entity.Property(e => e.Repnombres)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repnombres");
            entity.Property(e => e.Repnumdocu)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("repnumdocu");
            entity.Property(e => e.ServerEnvio).HasColumnName("server_envio");
            entity.Property(e => e.Telefono)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("telefono");
            entity.Property(e => e.Ubigeodir)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("ubigeodir");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
            entity.Property(e => e.UsuarioEnvio)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("usuario_envio");

            entity.HasOne(d => d.CodviaNavigation).WithMany(p => p.Plcfgempresas)
                .HasForeignKey(d => d.Codvia)
                .OnDelete(DeleteBehavior.SetNull);

            entity.HasOne(d => d.CodzonaNavigation).WithMany(p => p.Plcfgempresas)
                .HasForeignKey(d => d.Codzona)
                .OnDelete(DeleteBehavior.SetNull);

            entity.HasOne(d => d.GercoddciNavigation).WithMany(p => p.PlcfgempresaGercoddciNavigations)
                .HasForeignKey(d => d.Gercoddci)
                .OnDelete(DeleteBehavior.SetNull);

            entity.HasOne(d => d.RepcoddciNavigation).WithMany(p => p.PlcfgempresaRepcoddciNavigations).HasForeignKey(d => d.Repcoddci);
        });

        modelBuilder.Entity<Plclasplan>(entity =>
        {
            entity.HasKey(e => e.Codcls).HasName("PK_plclasplan_codcls");

            entity.ToTable("plclasplan", tb => tb.HasComment("Maestro de clase de planilla"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Clave)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("clave");
            entity.Property(e => e.Descls)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("descls");
            entity.Property(e => e.Estadocls)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocls");
            entity.Property(e => e.Fmtboleta)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fmtboleta");
            entity.Property(e => e.Fmtrecibo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fmtrecibo");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Horadiaria)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("horadiaria");
            entity.Property(e => e.Tipo)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("tipo");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plcodigoldn>(entity =>
        {
            entity.HasKey(e => e.Codldn).HasName("PK_plcodigoldn_codldn");

            entity.ToTable("plcodigoldn", tb => tb.HasComment("Maestro de código larga distancia nacional"));

            entity.Property(e => e.Codldn)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codldn");
            entity.Property(e => e.Desldn)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("desldn");
            entity.Property(e => e.Estadoldn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoldn");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plcomprobantect>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK_plcomprobantect_clspsnorden");

            entity.ToTable("plcomprobantect", tb => tb.HasComment("Maestro de Comprobantes de Cuarta categoria"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Fecemision).HasColumnName("fecemision");
            entity.Property(e => e.Fecpago).HasColumnName("fecpago");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Monto)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("monto");
            entity.Property(e => e.Numero)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("numero");
            entity.Property(e => e.Retencion)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("retencion");
            entity.Property(e => e.Serie)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("serie");
            entity.Property(e => e.Tipo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipo");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plcomprobantects)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .HasConstraintName("FK_plcomprobantect_plpersonal_clspsn");
        });

        modelBuilder.Entity<Plconceplanilla>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codcpc }).HasName("PK_plconceplanilla_clscpc");

            entity.ToTable("plconceplanilla", tb => tb.HasComment("Conceptos de calculo por clase planilla"));

            entity.HasIndex(e => e.Codcpc, "IX_plconceplanilla_codcpc");

            entity.HasIndex(e => e.Codsunat, "IX_plconceplanilla_codsunat");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Clasecpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("clasecpc");
            entity.Property(e => e.Codsunat)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codsunat");
            entity.Property(e => e.Defaultcpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("defaultcpc");
            entity.Property(e => e.Formulafun)
                .HasMaxLength(255)
                .IsUnicode(false)
                .HasColumnName("formulafun");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Imagenfun)
                .HasMaxLength(255)
                .IsUnicode(false)
                .HasColumnName("imagenfun");
            entity.Property(e => e.Impbolecpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("impbolecpc");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plconceplanillas)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.CodcpcNavigation).WithMany(p => p.Plconceplanillas)
                .HasForeignKey(d => d.Codcpc)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.CodsunatNavigation).WithMany(p => p.Plconceplanillas)
                .HasForeignKey(d => d.Codsunat)
                .OnDelete(DeleteBehavior.SetNull)
                .HasConstraintName("FK_plconceplanilla_plconcesunat");
        });

        modelBuilder.Entity<Plconceproceso>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codproce, e.Codcpc, e.Secuencia }).HasName("PK_plconceproceso_clsprocecpcsecuencia");

            entity.ToTable("plconceproceso");

            entity.HasIndex(e => e.Codcls, "IX_plconceproceso_codcpc");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codproce)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codproce");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Secuencia).HasColumnName("secuencia");
            entity.Property(e => e.Formulafun)
                .HasMaxLength(255)
                .IsUnicode(false)
                .HasColumnName("formulafun");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Codc).WithMany(p => p.Plconceprocesos)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plconceproceso_plconceplanilla_clscpc");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plconceprocesos)
                .HasForeignKey(d => new { d.Codcls, d.Codproce })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plconceproceso_plproceso_clsproce");
        });

        modelBuilder.Entity<Plconcepto>(entity =>
        {
            entity.HasKey(e => e.Codcpc).HasName("PK_plconcepto_codcpc");

            entity.ToTable("plconcepto", tb => tb.HasComment("Maestro de conceptos de calculo"));

            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Aliascpc)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("aliascpc");
            entity.Property(e => e.Descpc)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("descpc");
            entity.Property(e => e.Estadocpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocpc");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Obs)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("obs");
            entity.Property(e => e.Tipocpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipocpc");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plconcesunat>(entity =>
        {
            entity.HasKey(e => e.Codcon).HasName("PK_plconcesunat_codcon");

            entity.ToTable("plconcesunat", tb => tb.HasComment("Maestro de Conceptos Sunat"));

            entity.Property(e => e.Codcon)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcon");
            entity.Property(e => e.Csap)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("csap");
            entity.Property(e => e.Descon)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("descon");
            entity.Property(e => e.Estadocon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocon");
            entity.Property(e => e.Fcjmms)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fcjmms");
            entity.Property(e => e.Fdsa)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fdsa");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Ies)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("ies");
            entity.Property(e => e.Pensiones)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pensiones");
            entity.Property(e => e.Quinta)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("quinta");
            entity.Property(e => e.Reptp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("reptp");
            entity.Property(e => e.Sctr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sctr");
            entity.Property(e => e.Segagracui)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segagracui");
            entity.Property(e => e.Segreg)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segreg");
            entity.Property(e => e.Segregcbssp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segregcbssp");
            entity.Property(e => e.Segregpen)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segregpen");
            entity.Property(e => e.Senati)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("senati");
            entity.Property(e => e.Spp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("spp");
            entity.Property(e => e.Tipcon)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("tipcon");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plconditrabajo>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codcdt }).HasName("PK_plconditrabajo_clscdt");

            entity.ToTable("plconditrabajo", tb => tb.HasComment("Maestro Condicion de Trabajo"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codcdt)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcdt");
            entity.Property(e => e.Descdt)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("descdt");
            entity.Property(e => e.Estadocdt)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocdt");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plconditrabajos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plcontrato>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Numdocumen, e.Anyo, e.Mes, e.Dia }).HasName("PK_plcontrato_clspsndocu_anyomesdia");

            entity.ToTable("plcontrato", tb => tb.HasComment("Maestro de contrato de trabajo"));

            entity.HasIndex(e => e.Tipcon, "IX_plcontrato_tipcon");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Numdocumen)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("numdocumen");
            entity.Property(e => e.Anyo)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("anyo");
            entity.Property(e => e.Mes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mes");
            entity.Property(e => e.Dia)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("dia");
            entity.Property(e => e.Archivo)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("archivo");
            entity.Property(e => e.Estadocon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocon");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Observacion)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("observacion");
            entity.Property(e => e.Tipcon)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipcon");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.TipconNavigation).WithMany(p => p.Plcontratos).HasForeignKey(d => d.Tipcon);

            entity.HasOne(d => d.Cod).WithMany(p => p.Plcontratos)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plcontrato_plpersonal_clspsn");
        });

        modelBuilder.Entity<Plconven>(entity =>
        {
            entity.HasKey(e => e.Codctr).HasName("PK_plconven_codctr");

            entity.ToTable("plconven", tb => tb.HasComment("Maestro de Convenios evitar doble tributación"));

            entity.Property(e => e.Codctr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codctr");
            entity.Property(e => e.Desctr)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desctr");
            entity.Property(e => e.Estadoctr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoctr");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plctacenco>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codcco, e.Codsec, e.Codcpc, e.Orden }).HasName("PK_plctacencos_clsccosec_cpcorden");

            entity.ToTable("plctacencos", tb => tb.HasComment("Maestro de cuentas por centro de costo"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "IX_plctacencos_clscpc");

            entity.HasIndex(e => e.Codafp, "IX_plctacencos_codafp");

            entity.HasIndex(e => e.Codcco, "IX_plctacencos_codcco");

            entity.HasIndex(e => e.CodctaDebme, "IX_plctacencos_codcta_debme");

            entity.HasIndex(e => e.CodctaDebmn, "IX_plctacencos_codcta_debmn");

            entity.HasIndex(e => e.CodctaHabme, "IX_plctacencos_codcta_habme");

            entity.HasIndex(e => e.CodctaHabmn, "IX_plctacencos_codcta_habmn");

            entity.HasIndex(e => e.Codsec, "IX_plctacencos_codsec");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Codsec)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsec");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Codafp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codafp");
            entity.Property(e => e.CodctaDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debme");
            entity.Property(e => e.CodctaDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debmn");
            entity.Property(e => e.CodctaHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habme");
            entity.Property(e => e.CodctaHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habmn");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodafpNavigation).WithMany(p => p.Plctacencos)
                .HasForeignKey(d => d.Codafp)
                .HasConstraintName("FK_plctacencos_plentidadAFP_codafp");

            entity.HasOne(d => d.CodccoNavigation).WithMany(p => p.Plctacencos)
                .HasForeignKey(d => d.Codcco)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.CodctaDebmeNavigation).WithMany(p => p.PlctacencoCodctaDebmeNavigations).HasForeignKey(d => d.CodctaDebme);

            entity.HasOne(d => d.CodctaDebmnNavigation).WithMany(p => p.PlctacencoCodctaDebmnNavigations).HasForeignKey(d => d.CodctaDebmn);

            entity.HasOne(d => d.CodctaHabmeNavigation).WithMany(p => p.PlctacencoCodctaHabmeNavigations).HasForeignKey(d => d.CodctaHabme);

            entity.HasOne(d => d.CodctaHabmnNavigation).WithMany(p => p.PlctacencoCodctaHabmnNavigations).HasForeignKey(d => d.CodctaHabmn);

            entity.HasOne(d => d.CodsecNavigation).WithMany(p => p.Plctacencos)
                .HasForeignKey(d => d.Codsec)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.Codc).WithMany(p => p.Plctacencos)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plctacencos_plconceplanilla_clscpc");
        });

        modelBuilder.Entity<Plctapv>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codcco, e.Codsec, e.Orden }).HasName("PK_plctapvs_clsccosecorden");

            entity.ToTable("plctapvs", tb => tb.HasComment("Maestro de cuentas provisión"));

            entity.HasIndex(e => e.CodctagexDebme, "ID_ctapvs_codctagex_debme");

            entity.HasIndex(e => e.CodctagexDebmn, "ID_ctapvs_codctagex_debmn");

            entity.HasIndex(e => e.CodctagexHabme, "ID_ctapvs_codctagex_habme");

            entity.HasIndex(e => e.CodctagexHabmn, "ID_ctapvs_codctagex_habmn");

            entity.HasIndex(e => e.CodctagraDebme, "ID_ctapvs_codctagra_debme");

            entity.HasIndex(e => e.CodctagraDebmn, "ID_ctapvs_codctagra_debmn");

            entity.HasIndex(e => e.CodctagraHabme, "ID_ctapvs_codctagra_habme");

            entity.HasIndex(e => e.CodctagraHabmn, "ID_ctapvs_codctagra_habmn");

            entity.HasIndex(e => e.CodctavacDebme, "ID_ctapvs_codctavac_debme");

            entity.HasIndex(e => e.CodctavacDebmn, "ID_ctapvs_codctavac_debmn");

            entity.HasIndex(e => e.CodctavacHabme, "ID_ctapvs_codctavac_habme");

            entity.HasIndex(e => e.CodctavacHabmn, "ID_ctapvs_codctavac_habmn");

            entity.HasIndex(e => e.CodctavexDebme, "ID_ctapvs_codctavex_debme");

            entity.HasIndex(e => e.CodctavexDebmn, "ID_ctapvs_codctavex_debmn");

            entity.HasIndex(e => e.CodctavexHabme, "ID_ctapvs_codctavex_habme");

            entity.HasIndex(e => e.CodctavexHabmn, "ID_ctapvs_codctavex_habmn");

            entity.HasIndex(e => e.Codcco, "IX_plctapvs_codcco");

            entity.HasIndex(e => e.Codsec, "IX_plctapvs_seccion");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Codsec)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsec");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.CodctagexDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagex_debme");
            entity.Property(e => e.CodctagexDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagex_debmn");
            entity.Property(e => e.CodctagexHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagex_habme");
            entity.Property(e => e.CodctagexHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagex_habmn");
            entity.Property(e => e.CodctagraDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagra_debme");
            entity.Property(e => e.CodctagraDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagra_debmn");
            entity.Property(e => e.CodctagraHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagra_habme");
            entity.Property(e => e.CodctagraHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctagra_habmn");
            entity.Property(e => e.CodctavacDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavac_debme");
            entity.Property(e => e.CodctavacDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavac_debmn");
            entity.Property(e => e.CodctavacHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavac_habme");
            entity.Property(e => e.CodctavacHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavac_habmn");
            entity.Property(e => e.CodctavexDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavex_debme");
            entity.Property(e => e.CodctavexDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavex_debmn");
            entity.Property(e => e.CodctavexHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavex_habme");
            entity.Property(e => e.CodctavexHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codctavex_habmn");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodccoNavigation).WithMany(p => p.Plctapvs)
                .HasForeignKey(d => d.Codcco)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plctapvs)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.CodctagexDebmeNavigation).WithMany(p => p.PlctapvCodctagexDebmeNavigations).HasForeignKey(d => d.CodctagexDebme);

            entity.HasOne(d => d.CodctagexDebmnNavigation).WithMany(p => p.PlctapvCodctagexDebmnNavigations).HasForeignKey(d => d.CodctagexDebmn);

            entity.HasOne(d => d.CodctagexHabmeNavigation).WithMany(p => p.PlctapvCodctagexHabmeNavigations).HasForeignKey(d => d.CodctagexHabme);

            entity.HasOne(d => d.CodctagexHabmnNavigation).WithMany(p => p.PlctapvCodctagexHabmnNavigations).HasForeignKey(d => d.CodctagexHabmn);

            entity.HasOne(d => d.CodctagraDebmeNavigation).WithMany(p => p.PlctapvCodctagraDebmeNavigations).HasForeignKey(d => d.CodctagraDebme);

            entity.HasOne(d => d.CodctagraDebmnNavigation).WithMany(p => p.PlctapvCodctagraDebmnNavigations).HasForeignKey(d => d.CodctagraDebmn);

            entity.HasOne(d => d.CodctagraHabmeNavigation).WithMany(p => p.PlctapvCodctagraHabmeNavigations).HasForeignKey(d => d.CodctagraHabme);

            entity.HasOne(d => d.CodctagraHabmnNavigation).WithMany(p => p.PlctapvCodctagraHabmnNavigations).HasForeignKey(d => d.CodctagraHabmn);

            entity.HasOne(d => d.CodctavacDebmeNavigation).WithMany(p => p.PlctapvCodctavacDebmeNavigations).HasForeignKey(d => d.CodctavacDebme);

            entity.HasOne(d => d.CodctavacDebmnNavigation).WithMany(p => p.PlctapvCodctavacDebmnNavigations).HasForeignKey(d => d.CodctavacDebmn);

            entity.HasOne(d => d.CodctavacHabmeNavigation).WithMany(p => p.PlctapvCodctavacHabmeNavigations).HasForeignKey(d => d.CodctavacHabme);

            entity.HasOne(d => d.CodctavacHabmnNavigation).WithMany(p => p.PlctapvCodctavacHabmnNavigations).HasForeignKey(d => d.CodctavacHabmn);

            entity.HasOne(d => d.CodctavexDebmeNavigation).WithMany(p => p.PlctapvCodctavexDebmeNavigations).HasForeignKey(d => d.CodctavexDebme);

            entity.HasOne(d => d.CodctavexDebmnNavigation).WithMany(p => p.PlctapvCodctavexDebmnNavigations).HasForeignKey(d => d.CodctavexDebmn);

            entity.HasOne(d => d.CodctavexHabmeNavigation).WithMany(p => p.PlctapvCodctavexHabmeNavigations).HasForeignKey(d => d.CodctavexHabme);

            entity.HasOne(d => d.CodctavexHabmnNavigation).WithMany(p => p.PlctapvCodctavexHabmnNavigations).HasForeignKey(d => d.CodctavexHabmn);

            entity.HasOne(d => d.CodsecNavigation).WithMany(p => p.Plctapvs)
                .HasForeignKey(d => d.Codsec)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plctsmovimiento>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Pdocts, e.Subcts, e.Codpsn }).HasName("PK_plctsmovimiento_clspdocts_subctspsn");

            entity.ToTable("plctsmovimiento", tb => tb.HasComment("Movimiento CTS por personal"));

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "IX_plctsmovimiento_clspsn");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Pdocts)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("pdocts");
            entity.Property(e => e.Subcts)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("subcts");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Estadomov)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadomov");
            entity.Property(e => e.Fechacan)
                .HasColumnType("date")
                .HasColumnName("fechacan");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fechaven)
                .HasColumnType("date")
                .HasColumnName("fechaven");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Nrodeposito)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("nrodeposito");
            entity.Property(e => e.Numeroanos).HasColumnName("numeroanos");
            entity.Property(e => e.Numerodias).HasColumnName("numerodias");
            entity.Property(e => e.Numeromeses).HasColumnName("numeromeses");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Pdomes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdomes");
            entity.Property(e => e.Porinteres)
                .HasColumnType("decimal(5, 2)")
                .HasColumnName("porinteres");
            entity.Property(e => e.Tipocambio)
                .HasColumnType("decimal(6, 3)")
                .HasColumnName("tipocambio");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plctsmovimientos)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plctsmovimiento_plpersonal_clspsn");

            entity.HasOne(d => d.Plctsperiodosub).WithMany(p => p.Plctsmovimientos)
                .HasForeignKey(d => new { d.Codcls, d.Pdocts, d.Subcts })
                .HasConstraintName("FK_plctsmovimiento_plctsperiodosub_clspdocts_subcts");
        });

        modelBuilder.Entity<Plctsperiodo>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Pdocts }).HasName("PK_plctsperiodo_clspdocts");

            entity.ToTable("plctsperiodo", tb => tb.HasComment("Maestro de periodos de calculo CTS"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Pdocts)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("pdocts");
            entity.Property(e => e.Descricts)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descricts");
            entity.Property(e => e.Estadocts)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadocts");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plctsperiodos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plctsperiodosub>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Pdocts, e.Subcts }).HasName("PK_plctsperiodosub_clspdocts_subcts");

            entity.ToTable("plctsperiodosub", tb => tb.HasComment("Maestro sub periodos CTS"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Pdocts)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("pdocts");
            entity.Property(e => e.Subcts)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("subcts");
            entity.Property(e => e.Descrisub)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descrisub");
            entity.Property(e => e.Estadosub)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadosub");
            entity.Property(e => e.Fechacan)
                .HasColumnType("date")
                .HasColumnName("fechacan");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fechaven)
                .HasColumnType("date")
                .HasColumnName("fechaven");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Numeroanos).HasColumnName("numeroanos");
            entity.Property(e => e.Numerodias).HasColumnName("numerodias");
            entity.Property(e => e.Numeromeses).HasColumnName("numeromeses");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Pdomes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdomes");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Plctsperiodo).WithMany(p => p.Plctsperiodosubs)
                .HasForeignKey(d => new { d.Codcls, d.Pdocts })
                .HasConstraintName("FK_plctsperiodosub_plctsperiodo_clspdocts");
        });

        modelBuilder.Entity<Plctsresultado>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Pdocts, e.Subcts, e.Codpsn, e.Codcpc, e.Secuencia }).HasName("PK_plctsresultado_clspdoctssubcts_psncpcsecuencia");

            entity.ToTable("plctsresultado", tb => tb.HasComment("Resultado de Procesos de CTS"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "IX_plctsresultado_clscpc");

            entity.HasIndex(e => e.CodctaDebme, "IX_plctsresultado_codcta_debme");

            entity.HasIndex(e => e.CodctaDebmn, "IX_plctsresultado_codcta_debmn");

            entity.HasIndex(e => e.CodctaHabme, "IX_plctsresultado_codcta_habme");

            entity.HasIndex(e => e.CodctaHabmn, "IX_plctsresultado_codcta_habmn");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Pdocts)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("pdocts");
            entity.Property(e => e.Subcts)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("subcts");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Secuencia).HasColumnName("secuencia");
            entity.Property(e => e.CodctaDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debme");
            entity.Property(e => e.CodctaDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debmn");
            entity.Property(e => e.CodctaHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habme");
            entity.Property(e => e.CodctaHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habmn");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Impbolecpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("impbolecpc");
            entity.Property(e => e.ImporteMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_me");
            entity.Property(e => e.ImporteMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_mn");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Pdomes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdomes");
            entity.Property(e => e.Tipocpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipocpc");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodctaDebmeNavigation).WithMany(p => p.PlctsresultadoCodctaDebmeNavigations).HasForeignKey(d => d.CodctaDebme);

            entity.HasOne(d => d.CodctaDebmnNavigation).WithMany(p => p.PlctsresultadoCodctaDebmnNavigations).HasForeignKey(d => d.CodctaDebmn);

            entity.HasOne(d => d.CodctaHabmeNavigation).WithMany(p => p.PlctsresultadoCodctaHabmeNavigations).HasForeignKey(d => d.CodctaHabme);

            entity.HasOne(d => d.CodctaHabmnNavigation).WithMany(p => p.PlctsresultadoCodctaHabmnNavigations).HasForeignKey(d => d.CodctaHabmn);

            entity.HasOne(d => d.Codc).WithMany(p => p.Plctsresultados)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plctsresultado_plconceplanilla_clscpc");

            entity.HasOne(d => d.Plctsmovimiento).WithMany(p => p.Plctsresultados)
                .HasForeignKey(d => new { d.Codcls, d.Pdocts, d.Subcts, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_plctsresultado_plctsmovimiento_plclspdocts_subctspsn");
        });

        modelBuilder.Entity<Plcuentacte>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Numctacte, e.Numcuota }).HasName("PK_plcuentacte_clspsnnumctacuota");

            entity.ToTable("plcuentacte", tb => tb.HasComment("Cuenta corriente de adelanto o prestamo"));

            entity.HasIndex(e => new { e.Codcls, e.Codpdoprv }, "IX_plcuentacte__clspdoprv");

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "IX_plcuentacte_clscpc");

            entity.HasIndex(e => new { e.Codcls, e.Codpdocan }, "IX_plcuentacte_clspdocan");

            entity.HasIndex(e => e.Codbco, "IX_plcuentacte_codbco");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Numctacte)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("numctacte");
            entity.Property(e => e.Numcuota).HasColumnName("numcuota");
            entity.Property(e => e.AbonoMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("abono_me");
            entity.Property(e => e.AbonoMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("abono_mn");
            entity.Property(e => e.CargoMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("cargo_me");
            entity.Property(e => e.CargoMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("cargo_mn");
            entity.Property(e => e.Codbco)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbco");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Codpdocan)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdocan");
            entity.Property(e => e.Codpdoprv)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdoprv");
            entity.Property(e => e.Estadoctacte)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoctacte");
            entity.Property(e => e.Fectacte)
                .HasColumnType("date")
                .HasColumnName("fectacte");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Indchecar)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indchecar");
            entity.Property(e => e.Indgratifi)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indgratifi");
            entity.Property(e => e.Indprn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indprn");
            entity.Property(e => e.Numchecar)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("numchecar");
            entity.Property(e => e.Tpoctacte)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpoctacte");
            entity.Property(e => e.Tpodscto)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpodscto");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodbcoNavigation).WithMany(p => p.Plcuentactes).HasForeignKey(d => d.Codbco);

            entity.HasOne(d => d.Codc).WithMany(p => p.Plcuentactes)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .HasConstraintName("FK_plcuentacte_plconceplanilla_clscpc");

            entity.HasOne(d => d.Cod).WithMany(p => p.PlcuentacteCods)
                .HasForeignKey(d => new { d.Codcls, d.Codpdocan })
                .HasConstraintName("FK_plcuentacte_plperiodo_clspdocan");

            entity.HasOne(d => d.CodNavigation).WithMany(p => p.PlcuentacteCodNavigations)
                .HasForeignKey(d => new { d.Codcls, d.Codpdoprv })
                .HasConstraintName("FK_plcuentacte_plperiodo_clspdoprv");

            entity.HasOne(d => d.Cod1).WithMany(p => p.Plcuentactes)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .HasConstraintName("FK_plcuentacte_plpersonal_clspsn");
        });

        modelBuilder.Entity<Pldatoresultado>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo, e.Codpsn }).HasName("PK_pldatoresultado_clspdopsn");

            entity.ToTable("pldatoresultado", tb => tb.HasComment("Datos de proceso de calculo"));

            entity.HasIndex(e => new { e.Codcls, e.Codcdt }, "IX_pldatoresultado_clscdt");

            entity.HasIndex(e => new { e.Codcls, e.Codcgo }, "IX_pldatoresultado_clscgo");

            entity.HasIndex(e => e.Codafp, "IX_pldatoresultado_codafp");

            entity.HasIndex(e => e.Codcco, "IX_pldatoresultado_codcco");

            entity.HasIndex(e => e.Codeps, "IX_pldatoresultado_codeps");

            entity.HasIndex(e => e.Codsec, "IX_pldatoresultado_codsec");

            entity.HasIndex(e => e.Codubica, "IX_pldatoresultado_codubica");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codafp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codafp");
            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Codcdt)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcdt");
            entity.Property(e => e.Codcgo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcgo");
            entity.Property(e => e.Codeps)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codeps");
            entity.Property(e => e.Codsec)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsec");
            entity.Property(e => e.Codubica)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codubica");
            entity.Property(e => e.Estadopsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadopsn");
            entity.Property(e => e.Fecestado).HasColumnName("fecestado");
            entity.Property(e => e.Fecingreso)
                .HasColumnType("date")
                .HasColumnName("fecingreso");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Naciextrapsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("naciextrapsn");
            entity.Property(e => e.Regpension)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("regpension");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodafpNavigation).WithMany(p => p.Pldatoresultados).HasForeignKey(d => d.Codafp);

            entity.HasOne(d => d.CodccoNavigation).WithMany(p => p.Pldatoresultados).HasForeignKey(d => d.Codcco);

            entity.HasOne(d => d.CodepsNavigation).WithMany(p => p.Pldatoresultados).HasForeignKey(d => d.Codeps);

            entity.HasOne(d => d.CodsecNavigation).WithMany(p => p.Pldatoresultados).HasForeignKey(d => d.Codsec);

            entity.HasOne(d => d.CodubicaNavigation).WithMany(p => p.Pldatoresultados).HasForeignKey(d => d.Codubica);

            entity.HasOne(d => d.Codc).WithMany(p => p.Pldatoresultados)
                .HasForeignKey(d => new { d.Codcls, d.Codcdt })
                .HasConstraintName("FK_pldatoresultado_plconditrabajo_clscdt");

            entity.HasOne(d => d.CodcNavigation).WithMany(p => p.Pldatoresultados)
                .HasForeignKey(d => new { d.Codcls, d.Codcgo })
                .HasConstraintName("FK_pldatoresultado_plcargo_clscgo");

            entity.HasOne(d => d.Cod).WithMany(p => p.Pldatoresultados)
                .HasForeignKey(d => new { d.Codcls, d.Codpdo })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_pldatoresultado_plperiodo_clspdo");

            entity.HasOne(d => d.CodNavigation).WithMany(p => p.Pldatoresultados)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_pldatoresultado_plpersonal_clspsn");
        });

        modelBuilder.Entity<Pldetaboletum>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codboleta, e.Seccion, e.Dato, e.Tipodato, e.Fila, e.Columna }).HasName("PK_pldetaboleta_clsboletasecciondato_tipodatoficol");

            entity.ToTable("pldetaboleta", tb => tb.HasComment("Detalle de formatos de boleta de pago"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codboleta)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codboleta");
            entity.Property(e => e.Seccion)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("seccion");
            entity.Property(e => e.Dato)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("dato");
            entity.Property(e => e.Tipodato)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipodato");
            entity.Property(e => e.Fila).HasColumnName("fila");
            entity.Property(e => e.Columna).HasColumnName("columna");
            entity.Property(e => e.Desdato)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("desdato");
            entity.Property(e => e.Fontc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fontc");
            entity.Property(e => e.Fontn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fontn");
            entity.Property(e => e.Fonts)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("fonts");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Longitud).HasColumnName("longitud");
            entity.Property(e => e.Origen)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("origen");
            entity.Property(e => e.Sizefont)
                .HasColumnType("decimal(5, 2)")
                .HasColumnName("sizefont");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Cod).WithMany(p => p.Pldetaboleta)
                .HasForeignKey(d => new { d.Codcls, d.Codboleta })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_pldetaboleta_plboletapago_clsboleta");
        });

        modelBuilder.Entity<Pldetaplanilla>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpll, e.Fila, e.Columna, e.Codcpc }).HasName("PK__pldetapl__0C00508F30002040");

            entity.ToTable("pldetaplanilla", tb => tb.HasComment("Detalle Formato de planilla ministerio"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "ID_detaplanilla_clscpc");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpll)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codpll");
            entity.Property(e => e.Fila).HasColumnName("fila");
            entity.Property(e => e.Columna).HasColumnName("columna");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Codc).WithMany(p => p.Pldetaplanillas)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_pldetaplanilla_plconceplanilla");

            entity.HasOne(d => d.Plplanilla).WithMany(p => p.Pldetaplanillas)
                .HasForeignKey(d => new { d.Codcls, d.Codpll, d.Fila, d.Columna })
                .HasConstraintName("FK_pldetaplanilla_plplanilla_clspll_filcol");
        });

        modelBuilder.Entity<Pldetareporte>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codrpt, e.Orden }).HasName("PK_pldetareporte_clsrptorden");

            entity.ToTable("pldetareporte", tb => tb.HasComment("Detalle formato de generador de reporte"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codrpt)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codrpt");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Alias)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("alias");
            entity.Property(e => e.Descripcion)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descripcion");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Impreso)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("impreso");
            entity.Property(e => e.Longitud).HasColumnName("longitud");
            entity.Property(e => e.Nivel).HasColumnName("nivel");
            entity.Property(e => e.Signo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("signo");
            entity.Property(e => e.Tipo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasComment("A:cumulador, C:oncepto, D:ato")
                .HasColumnName("tipo");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Cod).WithMany(p => p.Pldetareportes)
                .HasForeignKey(d => new { d.Codcls, d.Codrpt })
                .HasConstraintName("FK_pldetareporte_plgenreporte_clsrpt");
        });

        modelBuilder.Entity<Pldocidentidad>(entity =>
        {
            entity.HasKey(e => e.Coddci).HasName("PK_pldocidentidad_coddci");

            entity.ToTable("pldocidentidad", tb => tb.HasComment("Maestro de documento de identidad"));

            entity.Property(e => e.Coddci)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("coddci");
            entity.Property(e => e.Codsunat)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("codsunat");
            entity.Property(e => e.Desdci)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desdci");
            entity.Property(e => e.Estadodci)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadodci");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Sigladci)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sigladci");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pldocvinfami>(entity =>
        {
            entity.HasKey(e => e.Coddvifa).HasName("PK__pldocvin__28BDD0627CF6B8EE");

            entity.ToTable("pldocvinfami", tb => tb.HasComment("Maestro Tipo Documento Vínculo Familiar"));

            entity.Property(e => e.Coddvifa)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("coddvifa");
            entity.Property(e => e.Desdvifa)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desdvifa");
            entity.Property(e => e.Estadotsu)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotsu");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pldstmonedum>(entity =>
        {
            entity.HasKey(e => new { e.Codmon, e.Valordmo }).HasName("PK_pldstmoneda_monvalordmo");

            entity.ToTable("pldstmoneda", tb => tb.HasComment("Maestro de billetes y monedas"));

            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Valordmo)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("valordmo");
            entity.Property(e => e.Desdmo)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desdmo");
            entity.Property(e => e.Estadodmo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadodmo");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plempleadore>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK_plempleadores_clspsnorden");

            entity.ToTable("plempleadores", tb => tb.HasComment("Maestro de Empleadores destaco Personal"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Razons)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("razons");
            entity.Property(e => e.Ruc)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("ruc");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plempresaseqde>(entity =>
        {
            entity.HasKey(e => e.Codeed).HasName("PK__plempres__47F9E309D4A99E03");

            entity.ToTable("plempresaseqdes", tb => tb.HasComment("Maestro Empresa Destaca o Desplaza Personal"));

            entity.Property(e => e.Codeed)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codeed");
            entity.Property(e => e.Deseed)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("deseed");
            entity.Property(e => e.Estadoeed)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoeed");
            entity.Property(e => e.Esteed)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("esteed");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Indeed)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indeed");
            entity.Property(e => e.Ruceed)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("ruceed");
            entity.Property(e => e.Taseed)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("taseed");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plempresasqde>(entity =>
        {
            entity.HasKey(e => e.Codeqd).HasName("PK__plempres__47FE028610FAB334");

            entity.ToTable("plempresasqdes", tb => tb.HasComment("Maestro Empresas que Destaca o Desplaza Personal"));

            entity.Property(e => e.Codeqd)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codeqd");
            entity.Property(e => e.Acteqd)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("acteqd");
            entity.Property(e => e.Deseqd)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("deseqd");
            entity.Property(e => e.Estadoeqd)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoeqd");
            entity.Property(e => e.FechafinEqd)
                .HasColumnType("date")
                .HasColumnName("fechafin_eqd");
            entity.Property(e => e.FechainiEqd)
                .HasColumnType("date")
                .HasColumnName("fechaini_eqd");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plempresasqmde>(entity =>
        {
            entity.HasKey(e => new { e.Codqmd, e.Desqmd }).HasName("PK__plempres__7F9F7EE2724E4098");

            entity.ToTable("plempresasqmdes", tb => tb.HasComment("Maestro Empresa que Desatacan o Desplazan Personal"));

            entity.Property(e => e.Codqmd)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codqmd");
            entity.Property(e => e.Desqmd)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desqmd");
            entity.Property(e => e.Actqmd)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("actqmd");
            entity.Property(e => e.Estadoqmd)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoqmd");
            entity.Property(e => e.FechafinQmd)
                .HasColumnType("date")
                .HasColumnName("fechafin_qmd");
            entity.Property(e => e.FechainiQmd)
                .HasColumnType("date")
                .HasColumnName("fechaini_qmd");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plentidadafp>(entity =>
        {
            entity.HasKey(e => e.Codafp).HasName("PK__plentida__86AECCB3C5C44503");

            entity.ToTable("plentidadafp", tb => tb.HasComment("Maestro de regimen pensionario - AFP"));

            entity.HasIndex(e => e.Codbco, "ID_entidadAFP_codbco");

            entity.Property(e => e.Codafp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codafp");
            entity.Property(e => e.Codbco)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbco");
            entity.Property(e => e.Codsunat)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("codsunat");
            entity.Property(e => e.Ctacteafp)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("ctacteafp");
            entity.Property(e => e.Ctactefondo)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("ctactefondo");
            entity.Property(e => e.Desafp)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("desafp");
            entity.Property(e => e.Desctacteafp)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desctacteafp");
            entity.Property(e => e.Desctactefondo)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desctactefondo");
            entity.Property(e => e.Estadoafp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoafp");
            entity.Property(e => e.Factor1)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factor1");
            entity.Property(e => e.Factor2)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factor2");
            entity.Property(e => e.Factor3)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factor3");
            entity.Property(e => e.Factor4)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factor4");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plentidadep>(entity =>
        {
            entity.HasKey(e => e.Codeps).HasName("PK__plentida__47FE0AD692F44995");

            entity.ToTable("plentidadeps", tb => tb.HasComment("Maestro de empresas prestadora de servicios"));

            entity.Property(e => e.Codeps)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codeps");
            entity.Property(e => e.Codsunat)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("codsunat");
            entity.Property(e => e.Deseps)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("deseps");
            entity.Property(e => e.Estadoeps)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoeps");
            entity.Property(e => e.Factoreps)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factoreps");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Ruceps)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("ruceps");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plescalaquintum>(entity =>
        {
            entity.HasKey(e => new { e.Pdoanyo, e.Orden }).HasName("PK__plescala__1F1A9E0C8495A391");

            entity.ToTable("plescalaquinta", tb => tb.HasComment("Configuracion escala de rango quinta categoria"));

            entity.Property(e => e.Pdoanyo)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoanyo");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Factor)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("factor");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Numerouit).HasColumnName("numerouit");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plestablecimiento>(entity =>
        {
            entity.HasKey(e => e.Codest).HasName("PK__plestabl__47FE74D71E7ECDE8");

            entity.ToTable("plestablecimiento", tb => tb.HasComment("Maestro Tipo de Establecimiento"));

            entity.Property(e => e.Codest)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codest");
            entity.Property(e => e.Desest)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desest");
            entity.Property(e => e.Estadoest)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoest");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plestablecimientopropio>(entity =>
        {
            entity.HasKey(e => e.Codepr).HasName("PK__plestabl__47FE0AE9A2C420BD");

            entity.ToTable("plestablecimientopropio", tb => tb.HasComment("Maestro de Establecimientos Propios"));

            entity.Property(e => e.Codepr)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codepr");
            entity.Property(e => e.Cdgepr)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("cdgepr");
            entity.Property(e => e.Desepr)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desepr");
            entity.Property(e => e.Estadoepr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoepr");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Indepr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indepr");
            entity.Property(e => e.Tasepr)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("tasepr");
            entity.Property(e => e.Tipepr)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipepr");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plestalaboral>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden, e.Ano, e.Mes }).HasName("PK__plestala__637A875DF0BF9FAE");

            entity.ToTable("plestalaboral", tb => tb.HasComment("Maestro Establecimiento Labora Trabajador"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Ano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("ano");
            entity.Property(e => e.Mes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mes");
            entity.Property(e => e.Codest)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codest");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Ruc)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("ruc");
            entity.Property(e => e.Tasa)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("tasa");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plestudio>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK__plestudi__C62CBD77DB878949");

            entity.ToTable("plestudios", tb => tb.HasComment("Maestro de estudio realizado"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Grado)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("grado");
            entity.Property(e => e.Institucion)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("institucion");
            entity.Property(e => e.Observacion)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("observacion");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plestudios)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_estudios_personal_clspsn");
        });

        modelBuilder.Entity<Plexpelaboral>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK__plexpela__C62CBD77BB52F079");

            entity.ToTable("plexpelaboral", tb => tb.HasComment("Maestro de experiencia laboral"));

            entity.HasIndex(e => new { e.Codcls, e.Codcgo }, "ID_expelaboral_clscgo");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Codcgo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcgo");
            entity.Property(e => e.Empresa)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("empresa");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Observacion)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("observacion");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Codc).WithMany(p => p.Plexpelaborals)
                .HasForeignKey(d => new { d.Codcls, d.Codcgo })
                .HasConstraintName("FK_expelaboral_cargo_clscgo");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plexpelaborals)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_expelaboral_personal_clspsn");
        });

        modelBuilder.Entity<Plfamiliare>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK__plfamili__C62CBD77D60EA863");

            entity.ToTable("plfamiliares", tb => tb.HasComment("Maestro Familiares Trabajador"));

            entity.HasIndex(e => e.Coddci, "ID_familiares_coddci");

            entity.HasIndex(e => e.Codvia, "ID_familiares_codvia");

            entity.HasIndex(e => e.Codzona, "ID_familiares_codzona");

            entity.HasIndex(e => e.Motivoina, "ID_familiares_motivoina");

            entity.HasIndex(e => e.Tipdocpaternidad, "ID_familiares_tipdocpaternidad");

            entity.HasIndex(e => e.Ubigeodom, "ID_familiares_ubigeodom");

            entity.HasIndex(e => e.Vinculo, "ID_familiares_vinculo");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Acrepaternidad)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("acrepaternidad");
            entity.Property(e => e.Apematerno)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("apematerno");
            entity.Property(e => e.Apepaterno)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("apepaterno");
            entity.Property(e => e.Cartamed)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cartamed");
            entity.Property(e => e.Certificadomed)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("certificadomed");
            entity.Property(e => e.Coddci)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("coddci");
            entity.Property(e => e.Codvia)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codvia");
            entity.Property(e => e.Codzona)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codzona");
            entity.Property(e => e.Domicilio)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("domicilio");
            entity.Property(e => e.Estadofam)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadofam");
            entity.Property(e => e.Fecalta).HasColumnName("fecalta");
            entity.Property(e => e.Fecbaja).HasColumnName("fecbaja");
            entity.Property(e => e.Fecnacimiento)
                .HasColumnType("date")
                .HasColumnName("fecnacimiento");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Incapacidad)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("incapacidad");
            entity.Property(e => e.Intedom)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("intedom");
            entity.Property(e => e.Motivoina)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("motivoina");
            entity.Property(e => e.Nombres)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nombres");
            entity.Property(e => e.Nomviadom)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nomviadom");
            entity.Property(e => e.Nomzonadom)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nomzonadom");
            entity.Property(e => e.Numdociden)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("numdociden");
            entity.Property(e => e.Numerdom)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("numerdom");
            entity.Property(e => e.Refedom)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("refedom");
            entity.Property(e => e.Sexofam)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sexofam");
            entity.Property(e => e.Tipdocpaternidad)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipdocpaternidad");
            entity.Property(e => e.Ubigeodom)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("ubigeodom");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
            entity.Property(e => e.Vinculo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("vinculo");
        });

        modelBuilder.Entity<Plgenreporte>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codrpt }).HasName("PK_plgenreporte_clsrpt");

            entity.ToTable("plgenreporte", tb => tb.HasComment("Cabecera de generador de reporte"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codrpt)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codrpt");
            entity.Property(e => e.Anchorpt).HasColumnName("anchorpt");
            entity.Property(e => e.Desrpt)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desrpt");
            entity.Property(e => e.Formarpt)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("formarpt");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Interlinea)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("interlinea");
            entity.Property(e => e.Pierpt)
                .HasMaxLength(255)
                .IsUnicode(false)
                .HasColumnName("pierpt");
            entity.Property(e => e.Titulorpt)
                .HasMaxLength(255)
                .IsUnicode(false)
                .HasColumnName("titulorpt");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plgenreportes)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plmodforma>(entity =>
        {
            entity.HasKey(e => e.Codmfo).HasName("PK__plmodfor__5DDEB929179A5B00");

            entity.ToTable("plmodforma", tb => tb.HasComment("Maestro Tipo de Modalidad Formativa Laboral"));

            entity.Property(e => e.Codmfo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmfo");
            entity.Property(e => e.Desmfo)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("desmfo");
            entity.Property(e => e.Estadomfo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadomfo");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plmotbajadh>(entity =>
        {
            entity.HasKey(e => e.Codbdh).HasName("PK__plmotbaj__40A868F79DA397B0");

            entity.ToTable("plmotbajadh", tb => tb.HasComment("Maestro Motivo de Baja de Derecho Habiente"));

            entity.Property(e => e.Codbdh)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbdh");
            entity.Property(e => e.Desbdh)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("desbdh");
            entity.Property(e => e.Estadobdh)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadobdh");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plmotfin>(entity =>
        {
            entity.HasKey(e => e.Codmof).HasName("PK__plmotfin__5DDEF1CF889F7973");

            entity.ToTable("plmotfin", tb => tb.HasComment("Maestro Motivo Fin Contrato o baja TRegistro"));

            entity.Property(e => e.Codmof)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmof");
            entity.Property(e => e.Desmof)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desmof");
            entity.Property(e => e.Estadomof)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadomof");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plnacionalidad>(entity =>
        {
            entity.HasKey(e => e.Codnac).HasName("PK__plnacion__47BB7FA430631B33");

            entity.ToTable("plnacionalidad", tb => tb.HasComment("Maestro Nacionalidad"));

            entity.HasIndex(e => e.Codpemi, "ID_nacionalidad_codpemi");

            entity.Property(e => e.Codnac)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codnac");
            entity.Property(e => e.Codpemi)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codpemi");
            entity.Property(e => e.Desnac)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("desnac");
            entity.Property(e => e.Estadonac)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadonac");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plniveducativo>(entity =>
        {
            entity.HasKey(e => e.Codniv).HasName("PK__plnivedu__47B8208B4DD1AA94");

            entity.ToTable("plniveducativo", tb => tb.HasComment("Maestro Situación Educativa"));

            entity.Property(e => e.Codniv)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codniv");
            entity.Property(e => e.Desniv)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("desniv");
            entity.Property(e => e.Estadoniv)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoniv");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plpaisemidocum>(entity =>
        {
            entity.HasKey(e => e.Codpemi).HasName("PK__plpaisem__9F17C1C3CEEB2034");

            entity.ToTable("plpaisemidocum", tb => tb.HasComment("Maestro Pais Emisor Documento"));

            entity.Property(e => e.Codpemi)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codpemi");
            entity.Property(e => e.Despemi)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("despemi");
            entity.Property(e => e.Estadopemi)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadopemi");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plperiodicidad>(entity =>
        {
            entity.HasKey(e => e.Codprd).HasName("PK__plperiod__473BF5A0D78D1A1B");

            entity.ToTable("plperiodicidad", tb => tb.HasComment("Maestro periodicidad de pago"));

            entity.Property(e => e.Codprd)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codprd");
            entity.Property(e => e.Desprd)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desprd");
            entity.Property(e => e.Estadoprd)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoprd");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plperiodo>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo }).HasName("PK_plperiodo_clspdo");

            entity.ToTable("plperiodo", tb => tb.HasComment("Maestro de periodos de pago"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Anopdo)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("anopdo");
            entity.Property(e => e.Despdo)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("despdo");
            entity.Property(e => e.Estadopdo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadopdo");
            entity.Property(e => e.Fechafin)
                .HasColumnType("date")
                .HasColumnName("fechafin");
            entity.Property(e => e.Fechaini)
                .HasColumnType("date")
                .HasColumnName("fechaini");
            entity.Property(e => e.Fechapago)
                .HasColumnType("date")
                .HasColumnName("fechapago");
            entity.Property(e => e.Fechaproceso)
                .HasColumnType("date")
                .HasColumnName("fechaproceso");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Mespdo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mespdo");
            entity.Property(e => e.Tipocambio)
                .HasColumnType("decimal(6, 3)")
                .HasColumnName("tipocambio");
            entity.Property(e => e.Tpopdo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tpopdo");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plperiodos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("PK_periodo_clasplan_codcls");
        });

        modelBuilder.Entity<Plpersonal>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn }).HasName("PK_plpersonal_clspsn");

            entity.ToTable("plpersonal", tb => tb.HasComment("Maestro de personal"));

            entity.HasIndex(e => e.Codbnkcts, "ID_personal_bnkcts");

            entity.HasIndex(e => e.Codbnkpago, "ID_personal_bnkpago");

            entity.HasIndex(e => e.Cmbcatocupacional, "ID_personal_catocupacion");

            entity.HasIndex(e => e.Cgoconfianza, "ID_personal_cgoconfianza");

            entity.HasIndex(e => new { e.Codcls, e.Variacpc }, "ID_personal_clsajustacpc");

            entity.HasIndex(e => new { e.Codcls, e.Netocpc }, "ID_personal_clsnetocpc");

            entity.HasIndex(e => e.Codafp, "ID_personal_codafp");

            entity.HasIndex(e => e.Codcco, "ID_personal_codcco");

            entity.HasIndex(e => e.Coddci, "ID_personal_coddci");

            entity.HasIndex(e => e.Codeps, "ID_personal_codeps");

            entity.HasIndex(e => e.Codldn, "ID_personal_codldn");

            entity.HasIndex(e => e.Codniv, "ID_personal_codniv");

            entity.HasIndex(e => e.Codpemi, "ID_personal_codpemi");

            entity.HasIndex(e => e.Codpfs, "ID_personal_codpfs");

            entity.HasIndex(e => e.Codsec, "ID_personal_codsec");

            entity.HasIndex(e => e.Codbcocts, "ID_plpersonal_bcocts");

            entity.HasIndex(e => new { e.Codcls, e.Codcgo }, "IX_plpersonal_clscgo");

            entity.HasIndex(e => e.Cmbtributacion, "IX_plpersonal_cmbtributacion");

            entity.HasIndex(e => e.Codbcopago, "IX_plpersonal_codbcopago");

            entity.HasIndex(e => e.Codtpt, "IX_plpersonal_codtpt");

            entity.HasIndex(e => e.Codubica, "IX_plpersonal_codubica");

            entity.HasIndex(e => e.Codvia, "IX_plpersonal_codvia");

            entity.HasIndex(e => e.Codzona, "IX_plpersonal_codzona");

            entity.HasIndex(e => e.Finperiodo, "IX_plpersonal_finperiodo");

            entity.HasIndex(e => e.Modformativa, "IX_plpersonal_modformativa");

            entity.HasIndex(e => e.Nacionalidad, "IX_plpersonal_nacionalidad");

            entity.HasIndex(e => e.Periodicidad, "IX_plpersonal_periodicidad");

            entity.HasIndex(e => e.Siteps, "IX_plpersonal_siteps");

            entity.HasIndex(e => e.Tippago, "IX_plpersonal_tippago");

            entity.HasIndex(e => e.Ubigeodir, "IX_plpersonal_ubigeodir");

            entity.HasIndex(e => e.Ubigeonac, "IX_plpersonal_ubigeonac");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Afilsindical)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("afilsindical");
            entity.Property(e => e.Afpmixta)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("afpmixta");
            entity.Property(e => e.Apematerno)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("apematerno");
            entity.Property(e => e.Apepaterno)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("apepaterno");
            entity.Property(e => e.Celular)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("celular");
            entity.Property(e => e.Cgoconfianza)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("cgoconfianza");
            entity.Property(e => e.Chk27252)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chk27252");
            entity.Property(e => e.ChkDis)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkDIS");
            entity.Property(e => e.ChkMax)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkMAX");
            entity.Property(e => e.ChkNoc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkNOC");
            entity.Property(e => e.ChkOiq)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkOIQ");
            entity.Property(e => e.ChkPe)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkPE");
            entity.Property(e => e.ChkQui)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkQUI");
            entity.Property(e => e.ChkReg)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkREG");
            entity.Property(e => e.ChkRl)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkRL");
            entity.Property(e => e.ChkSctrp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("chkSCTRP");
            entity.Property(e => e.Cmbcatocupacional)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("cmbcatocupacional");
            entity.Property(e => e.Cmbtributacion)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("cmbtributacion");
            entity.Property(e => e.Cobsctr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("cobsctr");
            entity.Property(e => e.Codacredor)
                .HasMaxLength(12)
                .IsUnicode(false)
                .HasColumnName("codacredor");
            entity.Property(e => e.Codafp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codafp");
            entity.Property(e => e.Codbcocts)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbcocts");
            entity.Property(e => e.Codbcopago)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbcopago");
            entity.Property(e => e.Codbnkcts)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbnkcts");
            entity.Property(e => e.Codbnkpago)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codbnkpago");
            entity.Property(e => e.Codcco)
                .HasMaxLength(9)
                .IsUnicode(false)
                .HasColumnName("codcco");
            entity.Property(e => e.Codcdt)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcdt");
            entity.Property(e => e.Codcgo)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcgo");
            entity.Property(e => e.Coddci)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("coddci");
            entity.Property(e => e.Coddeudor)
                .HasMaxLength(12)
                .IsUnicode(false)
                .HasColumnName("coddeudor");
            entity.Property(e => e.Codeps)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codeps");
            entity.Property(e => e.Codldn)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codldn");
            entity.Property(e => e.Codniv)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codniv");
            entity.Property(e => e.Codpemi)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codpemi");
            entity.Property(e => e.Codpfs)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("codpfs");
            entity.Property(e => e.Codsec)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsec");
            entity.Property(e => e.Codtpt)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtpt");
            entity.Property(e => e.Codubica)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codubica");
            entity.Property(e => e.Codvia)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codvia");
            entity.Property(e => e.Codzona)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codzona");
            entity.Property(e => e.Correoelect)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("correoelect");
            entity.Property(e => e.Ctsdeposito)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("ctsdeposito");
            entity.Property(e => e.Ctsdolar)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("ctsdolar");
            entity.Property(e => e.Cuentacts)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cuentacts");
            entity.Property(e => e.Cuentaibankcts)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cuentaibankcts");
            entity.Property(e => e.Cuentapago)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("cuentapago");
            entity.Property(e => e.Dctojudicial)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("dctojudicial");
            entity.Property(e => e.Essvida)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("essvida");
            entity.Property(e => e.Estadopsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadopsn");
            entity.Property(e => e.Estcivilpsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estcivilpsn");
            entity.Property(e => e.Fecbaja)
                .HasColumnType("date")
                .HasColumnName("fecbaja");
            entity.Property(e => e.Fecestado)
                .HasColumnType("date")
                .HasColumnName("fecestado");
            entity.Property(e => e.Fecingregpen)
                .HasColumnType("date")
                .HasColumnName("fecingregpen");
            entity.Property(e => e.Fecingreso)
                .HasColumnType("date")
                .HasColumnName("fecingreso");
            entity.Property(e => e.Fecnacimiento)
                .HasColumnType("date")
                .HasColumnName("fecnacimiento");
            entity.Property(e => e.Finperiodo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("finperiodo");
            entity.Property(e => e.Forprofesional)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("forprofesional");
            entity.Property(e => e.Fotopsn).HasColumnName("fotopsn");
            entity.Property(e => e.Fyhcre)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf)
                .HasColumnType("smalldatetime")
                .HasColumnName("fyhmdf");
            entity.Property(e => e.Imporemuneto)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("imporemuneto");
            entity.Property(e => e.Intedirec)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("intedirec");
            entity.Property(e => e.Interbankcts)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("interbankcts");
            entity.Property(e => e.Interbankpago)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("interbankpago");
            entity.Property(e => e.Jornadalaboral)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("jornadalaboral");
            entity.Property(e => e.Modformativa)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("modformativa");
            entity.Property(e => e.Naciextrapsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("naciextrapsn");
            entity.Property(e => e.Nacionalidad)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("nacionalidad");
            entity.Property(e => e.Netocpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("netocpc");
            entity.Property(e => e.Nombres)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("nombres");
            entity.Property(e => e.Nomviadirec)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nomviadirec");
            entity.Property(e => e.Nomzondirec)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nomzondirec");
            entity.Property(e => e.Nroessalud)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("nroessalud");
            entity.Property(e => e.Numdepen).HasColumnName("numdepen");
            entity.Property(e => e.Numdociden)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("numdociden");
            entity.Property(e => e.Numdocmil)
                .HasMaxLength(12)
                .IsUnicode(false)
                .HasColumnName("numdocmil");
            entity.Property(e => e.Numerdirec)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("numerdirec");
            entity.Property(e => e.Numeroafp)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("numeroafp");
            entity.Property(e => e.Numhijo).HasColumnName("numhijo");
            entity.Property(e => e.Pagodolar)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pagodolar");
            entity.Property(e => e.Periodicidad)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("periodicidad");
            entity.Property(e => e.Pordsctojudi)
                .HasColumnType("decimal(6, 2)")
                .HasColumnName("pordsctojudi");
            entity.Property(e => e.Refedirec)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("refedirec");
            entity.Property(e => e.Regpension)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("regpension");
            entity.Property(e => e.Reingreso)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("reingreso");
            entity.Property(e => e.Remimprecisa)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("remimprecisa");
            entity.Property(e => e.Remintegralcts)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("remintegralcts");
            entity.Property(e => e.Remintegralgrati)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("remintegralgrati");
            entity.Property(e => e.Remintegralvaca)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("remintegralvaca");
            entity.Property(e => e.Remuneta)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("remuneta");
            entity.Property(e => e.Resfamiliar)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("resfamiliar");
            entity.Property(e => e.Segmedico)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("segmedico");
            entity.Property(e => e.Sexopsn)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sexopsn");
            entity.Property(e => e.Siteps)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("siteps");
            entity.Property(e => e.Telefono)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("telefono");
            entity.Property(e => e.Tippago)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tippago");
            entity.Property(e => e.Ubigeodir)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("ubigeodir");
            entity.Property(e => e.Ubigeonac)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("ubigeonac");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
            entity.Property(e => e.Variacpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("variacpc");

            entity.HasOne(d => d.CmbcatocupacionalNavigation).WithMany(p => p.Plpersonals)
                .HasForeignKey(d => d.Cmbcatocupacional)
                .HasConstraintName("FK_personal_catocu_catocupacion");

            entity.HasOne(d => d.CmbtributacionNavigation).WithMany(p => p.Plpersonals)
                .HasForeignKey(d => d.Cmbtributacion)
                .OnDelete(DeleteBehavior.SetNull);

            entity.HasOne(d => d.CodbcoctsNavigation).WithMany(p => p.PlpersonalCodbcoctsNavigations)
                .HasForeignKey(d => d.Codbcocts)
                .HasConstraintName("FK_personal_banco_bcocts");

            entity.HasOne(d => d.CodbcopagoNavigation).WithMany(p => p.PlpersonalCodbcopagoNavigations)
                .HasForeignKey(d => d.Codbcopago)
                .HasConstraintName("FK_personal_banco_bcopago");

            entity.HasOne(d => d.CodbnkctsNavigation).WithMany(p => p.PlpersonalCodbnkctsNavigations)
                .HasForeignKey(d => d.Codbnkcts)
                .HasConstraintName("FK_personal_banco_bnkcts");

            entity.HasOne(d => d.CodbnkpagoNavigation).WithMany(p => p.PlpersonalCodbnkpagoNavigations)
                .HasForeignKey(d => d.Codbnkpago)
                .HasConstraintName("FK_personal_banco_bnkpago");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plpersonals)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_personal_clasplan_codcls");

            entity.HasOne(d => d.Codc).WithMany(p => p.Plpersonals)
                .HasForeignKey(d => new { d.Codcls, d.Codcgo })
                .HasConstraintName("FK_plpersonal_plcargo_clscargo");
        });

        modelBuilder.Entity<Plplanilla>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpll, e.Fila, e.Columna }).HasName("PK_planilla_clspllfilacol");

            entity.ToTable("plplanilla", tb => tb.HasComment("Formato de planilla ministerio"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpll)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codpll");
            entity.Property(e => e.Fila).HasColumnName("fila");
            entity.Property(e => e.Columna).HasColumnName("columna");
            entity.Property(e => e.Alias)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("alias");
            entity.Property(e => e.Descripcion)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("descripcion");
            entity.Property(e => e.Despll)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("despll");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Imprimecab)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("imprimecab");
            entity.Property(e => e.Longitud).HasColumnName("longitud");
            entity.Property(e => e.Posicion).HasColumnName("posicion");
            entity.Property(e => e.Posipapel)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("posipapel");
            entity.Property(e => e.Sizefont)
                .HasColumnType("decimal(5, 2)")
                .HasColumnName("sizefont");
            entity.Property(e => e.Sizepapel)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sizepapel");
            entity.Property(e => e.Subrayado)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("subrayado");
            entity.Property(e => e.Tipo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasComment("C:oncepto, D:ato")
                .HasColumnName("tipo");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plplanillas)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plproceso>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codproce }).HasName("PK_plproceso_clsproce");

            entity.ToTable("plproceso", tb => tb.HasComment("Maestro de procesos de calculo"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codproce)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codproce");
            entity.Property(e => e.Desproce)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desproce");
            entity.Property(e => e.Estadoproce)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoproce");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodclsNavigation).WithMany(p => p.Plprocesos)
                .HasForeignKey(d => d.Codcls)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Plprofesion>(entity =>
        {
            entity.HasKey(e => e.Codpfs).HasName("PK__plprofes__473BD5DEDD789917");

            entity.ToTable("plprofesion", tb => tb.HasComment("Maestro de profesiones u ocupaciones"));

            entity.Property(e => e.Codpfs)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("codpfs");
            entity.Property(e => e.Cateobrpfs)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("cateobrpfs");
            entity.Property(e => e.Despfs)
                .HasMaxLength(100)
                .IsUnicode(false)
                .HasColumnName("despfs");
            entity.Property(e => e.Estadopfs)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadopfs");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plremudefa>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Codcpc }).HasName("PK__plremude__014D9C14380D9195");

            entity.ToTable("plremudefa", tb => tb.HasComment("Remuneración Default Trabajador"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "ID_remudefa_conceplanilla_clscpc");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Imporemune)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("imporemune");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Codc).WithMany(p => p.Plremudefas)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_remudefa_conceplanilla_clscpc");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plremudefas)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_remudefa_personal_clspsn");
        });

        modelBuilder.Entity<Plremuexce>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo, e.Codpsn, e.Codcpc }).HasName("PK__plremuex__E01ED20871AC3DB0");

            entity.ToTable("plremuexce", tb => tb.HasComment("Remuneraciones o eventos excepcionales"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "ID_remuexce_clscpc");

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "ID_remuexce_clspsn");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Imporemune)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("imporemune");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.Codc).WithMany(p => p.Plremuexces)
                .HasForeignKey(d => new { d.Codcls, d.Codcpc })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_remuexce_conceplanilla_clscpc");

            entity.HasOne(d => d.Cod).WithMany(p => p.Plremuexces)
                .HasForeignKey(d => new { d.Codcls, d.Codpdo })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_remuexce_periodo_clspdo");

            entity.HasOne(d => d.CodNavigation).WithMany(p => p.Plremuexces)
                .HasForeignKey(d => new { d.Codcls, d.Codpsn })
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_remuexce_personal_clspsn");
        });

        modelBuilder.Entity<Plresultado>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpdo, e.Codproce, e.Codpsn, e.Codcpc, e.Secuencia }).HasName("PK_plresultado_clspdoproce_psncpcsecuencia");

            entity.ToTable("plresultado", tb => tb.HasComment("Resultado proceso calculo planilla"));

            entity.HasIndex(e => new { e.Codcls, e.Codcpc }, "IX_plresultado_clscpc");

            entity.HasIndex(e => new { e.Codcls, e.Codpdo, e.Codpsn }, "IX_plresultado_clspdopsn");

            entity.HasIndex(e => new { e.Codcls, e.Codproce }, "IX_plresultado_clsproce");

            entity.HasIndex(e => new { e.Codcls, e.Codproce, e.Codcpc, e.Secuencia }, "IX_plresultado_clsproce_cpcsecuencia");

            entity.HasIndex(e => new { e.Codcls, e.CodprocePdo }, "IX_plresultado_clsproce_pdo");

            entity.HasIndex(e => new { e.Codcls, e.Codpsn }, "IX_plresultado_clspsn");

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpdo)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codpdo");
            entity.Property(e => e.Codproce)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codproce");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Codcpc)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codcpc");
            entity.Property(e => e.Secuencia).HasColumnName("secuencia");
            entity.Property(e => e.CodctaDebme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debme");
            entity.Property(e => e.CodctaDebmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_debmn");
            entity.Property(e => e.CodctaHabme)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habme");
            entity.Property(e => e.CodctaHabmn)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codcta_habmn");
            entity.Property(e => e.Codmon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codmon");
            entity.Property(e => e.CodprocePdo)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codproce_pdo");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Impbolecpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("impbolecpc");
            entity.Property(e => e.ImporteMe)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_me");
            entity.Property(e => e.ImporteMn)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("importe_mn");
            entity.Property(e => e.Pdoano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("pdoano");
            entity.Property(e => e.Pdomes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdomes");
            entity.Property(e => e.Tipocpc)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipocpc");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plseccion>(entity =>
        {
            entity.HasKey(e => e.Codsec).HasName("PK__plseccio__5C4967FDF7BDDEE6");

            entity.ToTable("plseccion", tb => tb.HasComment("Maestro de secciones de personal"));

            entity.Property(e => e.Codsec)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsec");
            entity.Property(e => e.Codintersec)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codintersec");
            entity.Property(e => e.Dessec)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("dessec");
            entity.Property(e => e.Estadosec)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadosec");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plsitrapen>(entity =>
        {
            entity.HasKey(e => e.Codstp).HasName("PK__plsitrap__5C5E236D00CD510A");

            entity.ToTable("plsitrapen", tb => tb.HasComment("Maestro Situacion Trabajador Pensionista"));

            entity.Property(e => e.Codstp)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codstp");
            entity.Property(e => e.Desstp)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("desstp");
            entity.Property(e => e.Estadostp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadostp");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plsituespecial>(entity =>
        {
            entity.HasKey(e => e.Codsie).HasName("PK__plsitues__5C5E3B36660C9EC1");

            entity.ToTable("plsituespecial", tb => tb.HasComment("Maestro Situación Especial"));

            entity.Property(e => e.Codsie)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsie");
            entity.Property(e => e.Dessie)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("dessie");
            entity.Property(e => e.Estadosie)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadosie");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plsuspensionct>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK__plsuspen__C62CBD776F29C18C");

            entity.ToTable("plsuspensionct", tb => tb.HasComment("Maestro de Suspension de Cuarta Categoria"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Ejercicio)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("ejercicio");
            entity.Property(e => e.Fecha).HasColumnName("fecha");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Medio)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("medio");
            entity.Property(e => e.Numero)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("numero");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltercero>(entity =>
        {
            entity.HasKey(e => new { e.Codcls, e.Codpsn, e.Orden }).HasName("PK__pltercer__C62CBD77F30F7121");

            entity.ToTable("plterceros", tb => tb.HasComment("Maestro Personal de Terceros"));

            entity.Property(e => e.Codcls)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codcls");
            entity.Property(e => e.Codpsn)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("codpsn");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Ano)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("ano");
            entity.Property(e => e.Codest)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codest");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Importe)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("importe");
            entity.Property(e => e.Mes)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mes");
            entity.Property(e => e.Ruc)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("ruc");
            entity.Property(e => e.Sctrp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sctrp");
            entity.Property(e => e.Sctrs)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sctrs");
            entity.Property(e => e.Tasa)
                .HasColumnType("decimal(10, 2)")
                .HasColumnName("tasa");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltipcom>(entity =>
        {
            entity.HasKey(e => e.Codtic).HasName("PK__pltipcom__40183C9F35FEB83E");

            entity.ToTable("pltipcom", tb => tb.HasComment("Maestro Tipo Comprobante Servicio"));

            entity.Property(e => e.Codtic)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtic");
            entity.Property(e => e.Destic)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("destic");
            entity.Property(e => e.Estadotic)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotic");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltipcontrato>(entity =>
        {
            entity.HasKey(e => e.Codtco).HasName("PK__pltipcon__401BEA589D6ED191");

            entity.ToTable("pltipcontrato", tb => tb.HasComment("Maestro Tipo de Contrato de Trabajo"));

            entity.Property(e => e.Codtco)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtco");
            entity.Property(e => e.Destco)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("destco");
            entity.Property(e => e.Estadotco)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotco");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltipovium>(entity =>
        {
            entity.HasKey(e => e.Codvia).HasName("PK__pltipovi__5D986088BB465534");

            entity.ToTable("pltipovia", tb => tb.HasComment("Maestro de tipo de via"));

            entity.Property(e => e.Codvia)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codvia");
            entity.Property(e => e.Abrevia)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("abrevia");
            entity.Property(e => e.Desvia)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desvia");
            entity.Property(e => e.Estadovia)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadovia");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltipozona>(entity =>
        {
            entity.HasKey(e => e.Codzona).HasName("PK__pltipozo__9D9E0D98573B8C59");

            entity.ToTable("pltipozona", tb => tb.HasComment("Maestro de tipo de zona"));

            entity.Property(e => e.Codzona)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codzona");
            entity.Property(e => e.Abrezona)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("abrezona");
            entity.Property(e => e.Deszona)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("deszona");
            entity.Property(e => e.Estadozona)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadozona");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltippago>(entity =>
        {
            entity.HasKey(e => e.Codtip).HasName("PK__pltippag__40183CA8D5D1431A");

            entity.ToTable("pltippago", tb => tb.HasComment("Maestro Tipo de Pago"));

            entity.Property(e => e.Codtip)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtip");
            entity.Property(e => e.Destip)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("destip");
            entity.Property(e => e.Estadotip)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotip");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltipsusp>(entity =>
        {
            entity.HasKey(e => e.Codtsu).HasName("PK_pltipsusp_codtsu");

            entity.ToTable("pltipsusp", tb => tb.HasComment("Maestro Tipo Suspension Laboral"));

            entity.Property(e => e.Codtsu)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtsu");
            entity.Property(e => e.Destsu)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("destsu");
            entity.Property(e => e.Estadotsu)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotsu");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Pltpotrabajador>(entity =>
        {
            entity.HasKey(e => e.Codtpt).HasName("PK__pltpotra__40198B87269A638D");

            entity.ToTable("pltpotrabajador", tb => tb.HasComment("Maestro de tipo de trabajador"));

            entity.Property(e => e.Codtpt)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtpt");
            entity.Property(e => e.Destpt)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("destpt");
            entity.Property(e => e.Estadotpt)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadotpt");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plubicacion>(entity =>
        {
            entity.HasKey(e => e.Codubica).HasName("PK__plubicac__65234D5702483BE4");

            entity.ToTable("plubicacion", tb => tb.HasComment("Maestro de ubicacion o localidad"));

            entity.Property(e => e.Codubica)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codubica");
            entity.Property(e => e.Codinterubica)
                .HasMaxLength(15)
                .IsUnicode(false)
                .HasColumnName("codinterubica");
            entity.Property(e => e.Desubica)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("desubica");
            entity.Property(e => e.Estadoubica)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadoubica");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Plvarfunc>(entity =>
        {
            entity.HasKey(e => new { e.Tipo, e.Codigo, e.Nombre }).HasName("PK__plvarfun__458463D42A6CD8F5");

            entity.ToTable("plvarfunc", tb => tb.HasComment("Maestro de Variable y Funcion de Cálculo"));

            entity.Property(e => e.Tipo)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("tipo");
            entity.Property(e => e.Codigo)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("codigo");
            entity.Property(e => e.Nombre)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("nombre");
            entity.Property(e => e.Descripcion)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("descripcion");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Orden).HasColumnName("orden");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
            entity.Property(e => e.Valor)
                .HasMaxLength(200)
                .IsUnicode(false)
                .HasColumnName("valor");
        });

        modelBuilder.Entity<Plvinfami>(entity =>
        {
            entity.HasKey(e => e.Codvfa).HasName("PK__plvinfam__5D98582AEFC927A8");

            entity.ToTable("plvinfami", tb => tb.HasComment("Maestro Vinculo Familiar"));

            entity.Property(e => e.Codvfa)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codvfa");
            entity.Property(e => e.Desvfa)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desvfa");
            entity.Property(e => e.Estadovfa)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estadovfa");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Rangoimpresion>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("rangoimpresion", tb => tb.HasComment("Rango de impresiones"));

            entity.HasIndex(e => new { e.Proceso, e.Valor }, "ID_rngimp");

            entity.Property(e => e.Fyhcre)
                .HasMaxLength(19)
                .IsUnicode(false)
                .HasColumnName("fyhcre");
            entity.Property(e => e.Proceso)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("proceso");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Valor)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("valor");
        });

        modelBuilder.Entity<Tgtcb>(entity =>
        {
            entity.HasKey(e => e.Fehtcb).HasName("PK__tgtcb__D51D382EFF0B64E9");

            entity.ToTable("tgtcb", tb => tb.HasComment("Maestro de tipo de cambio"));

            entity.Property(e => e.Fehtcb)
                .HasColumnType("date")
                .HasColumnName("fehtcb");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.ImptcbCpr).HasColumnName("imptcb_cpr");
            entity.Property(e => e.ImptcbVta).HasColumnName("imptcb_vta");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
