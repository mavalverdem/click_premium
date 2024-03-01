using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace ClickPremium.Infrastructure.Models.cfg;

public partial class ClickpremiumcfgContext : DbContext
{
    public ClickpremiumcfgContext()
    {
    }

    public ClickpremiumcfgContext(DbContextOptions<ClickpremiumcfgContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Pldocvinfami> Pldocvinfamis { get; set; }

    public virtual DbSet<Sgmdl> Sgmdls { get; set; }

    public virtual DbSet<Sgpm> Sgpms { get; set; }

    public virtual DbSet<Sgusr> Sgusrs { get; set; }

    public virtual DbSet<Tgctrobli> Tgctroblis { get; set; }

    public virtual DbSet<Tgctroblidet> Tgctroblidets { get; set; }

    public virtual DbSet<Tgemp> Tgemps { get; set; }

    public virtual DbSet<Tgsunat> Tgsunats { get; set; }

    public virtual DbSet<Tgubigeo> Tgubigeos { get; set; }

//     protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
// #warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
//         => optionsBuilder.UseSqlServer("Server=localhost;Database=CLICKPREMIUMCFG;User Id=sa;Password=P@ssw0rd@DB1;TrustServerCertificate=true");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Pldocvinfami>(entity =>
        {
            entity.HasKey(e => e.Coddvifa).HasName("PK__pldocvin__28BDD062FCF63685");

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

        modelBuilder.Entity<Sgmdl>(entity =>
        {
            entity.HasKey(e => new { e.Codsis, e.Opcion, e.Orden, e.Codmdl }).HasName("PK_sgmdl_sisopcordmdl");

            entity.ToTable("sgmdl", tb => tb.HasComment("Modulos u opciones del sistema"));

            entity.HasIndex(e => new { e.Codsis, e.Codmdl }, "IX_sgmdl_sismdl");

            entity.Property(e => e.Codsis)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsis");
            entity.Property(e => e.Opcion)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("opcion");
            entity.Property(e => e.Orden)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("orden");
            entity.Property(e => e.Codmdl)
                .HasMaxLength(16)
                .IsUnicode(false)
                .HasColumnName("codmdl");
            entity.Property(e => e.Detmdl)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("detmdl");
            entity.Property(e => e.Detmdlx)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("detmdlx");
            entity.Property(e => e.Estmdl)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estmdl");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Nommdl)
                .HasMaxLength(16)
                .IsUnicode(false)
                .HasColumnName("nommdl");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Sgpm>(entity =>
        {
            entity.HasKey(e => new { e.Codusr, e.Codemp, e.Codsis, e.Codmdl }).HasName("PK_sgpms_usrempsismdl");

            entity.ToTable("sgpms", tb => tb.HasComment("Permisos por usuario"));

            entity.HasIndex(e => new { e.Codsis, e.Codmdl }, "IX_sgpms_sismdl");

            entity.Property(e => e.Codusr)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("codusr");
            entity.Property(e => e.Codemp)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codemp");
            entity.Property(e => e.Codsis)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsis");
            entity.Property(e => e.Codmdl)
                .HasMaxLength(16)
                .IsUnicode(false)
                .HasColumnName("codmdl");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Indpms01).HasColumnName("indpms01");
            entity.Property(e => e.Indpms02).HasColumnName("indpms02");
            entity.Property(e => e.Indpms03).HasColumnName("indpms03");
            entity.Property(e => e.Indpms04).HasColumnName("indpms04");
            entity.Property(e => e.Indpms05).HasColumnName("indpms05");
            entity.Property(e => e.Indpms06).HasColumnName("indpms06");
            entity.Property(e => e.Indpms07).HasColumnName("indpms07");
            entity.Property(e => e.Indpms08).HasColumnName("indpms08");
            entity.Property(e => e.Indpms09).HasColumnName("indpms09");
            entity.Property(e => e.Indpms10).HasColumnName("indpms10");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodempNavigation).WithMany(p => p.Sgpms).HasForeignKey(d => d.Codemp);

            entity.HasOne(d => d.CodusrNavigation).WithMany(p => p.Sgpms)
                .HasForeignKey(d => d.Codusr)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Sgusr>(entity =>
        {
            entity.HasKey(e => e.Codusr).HasName("PK_sgusr_codusr");

            entity.ToTable("sgusr", tb => tb.HasComment("Maestro de usuarios"));

            entity.HasIndex(e => e.Empusr, "IX_sgusr_empusr");

            entity.Property(e => e.Codusr)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("codusr");
            entity.Property(e => e.Abvusr)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("abvusr");
            entity.Property(e => e.Anousr)
                .HasMaxLength(4)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("anousr");
            entity.Property(e => e.Clausr)
                .HasMaxLength(10)
                .IsUnicode(false)
                .HasColumnName("clausr");
            entity.Property(e => e.Empusr)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("empusr");
            entity.Property(e => e.Estusr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estusr");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Mesusr)
                .HasMaxLength(2)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("mesusr");
            entity.Property(e => e.Nomusr)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("nomusr");
            entity.Property(e => e.Nvlusr)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("nvlusr");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.EmpusrNavigation).WithMany(p => p.Sgusrs)
                .HasForeignKey(d => d.Empusr)
                .OnDelete(DeleteBehavior.Cascade);
        });

        modelBuilder.Entity<Tgctrobli>(entity =>
        {
            entity.HasKey(e => e.Pdotribu).HasName("PK_tgctrobli_pdotribu");

            entity.ToTable("tgctrobli", tb => tb.HasComment("Control de Obligaciones sunat - cabecera"));

            entity.Property(e => e.Pdotribu)
                .HasMaxLength(6)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdotribu");
            entity.Property(e => e.Buencontri)
                .HasColumnType("date")
                .HasColumnName("buencontri");
            entity.Property(e => e.FecVence0)
                .HasColumnType("date")
                .HasColumnName("fecVence0");
            entity.Property(e => e.FecVence1)
                .HasColumnType("date")
                .HasColumnName("fecVence1");
            entity.Property(e => e.FecVence2)
                .HasColumnType("date")
                .HasColumnName("fecVence2");
            entity.Property(e => e.FecVence3)
                .HasColumnType("date")
                .HasColumnName("fecVence3");
            entity.Property(e => e.FecVence4)
                .HasColumnType("date")
                .HasColumnName("fecVence4");
            entity.Property(e => e.FecVence5)
                .HasColumnType("date")
                .HasColumnName("fecVence5");
            entity.Property(e => e.FecVence6)
                .HasColumnType("date")
                .HasColumnName("fecVence6");
            entity.Property(e => e.FecVence7)
                .HasColumnType("date")
                .HasColumnName("fecVence7");
            entity.Property(e => e.FecVence8)
                .HasColumnType("date")
                .HasColumnName("fecVence8");
            entity.Property(e => e.FecVence9)
                .HasColumnType("date")
                .HasColumnName("fecVence9");
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

        modelBuilder.Entity<Tgctroblidet>(entity =>
        {
            entity.HasKey(e => new { e.Codemp, e.Pdotribu }).HasName("PK_tgctroblidet_emptribu");

            entity.ToTable("tgctroblidet", tb => tb.HasComment("Control de Obligaciones sunat - detalle"));

            entity.Property(e => e.Codemp)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codemp");
            entity.Property(e => e.Pdotribu)
                .HasMaxLength(6)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("pdotribu");
            entity.Property(e => e.Coddeclar)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("coddeclar");
            entity.Property(e => e.Fpresenta)
                .HasColumnType("date")
                .HasColumnName("fpresenta");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Nroconsta)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("nroconsta");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");

            entity.HasOne(d => d.CodempNavigation).WithMany(p => p.Tgctroblidets).HasForeignKey(d => d.Codemp);

            entity.HasOne(d => d.PdotribuNavigation).WithMany(p => p.Tgctroblidets).HasForeignKey(d => d.Pdotribu);
        });

        modelBuilder.Entity<Tgemp>(entity =>
        {
            entity.HasKey(e => e.Codemp).HasName("PK_tgemp_codemp");

            entity.ToTable("tgemp", tb => tb.HasComment("Maestro de empresas"));

            entity.Property(e => e.Codemp)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codemp");
            entity.Property(e => e.Actividad)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("actividad");
            entity.Property(e => e.Actividademp)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("actividademp");
            entity.Property(e => e.Buencontri)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("buencontri");
            entity.Property(e => e.CodctaPer)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codcta_per");
            entity.Property(e => e.CodctaRet)
                .HasMaxLength(8)
                .IsUnicode(false)
                .HasColumnName("codcta_ret");
            entity.Property(e => e.Conapematerno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("conapematerno");
            entity.Property(e => e.Conapepaterno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("conapepaterno");
            entity.Property(e => e.Condocumento)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("condocumento");
            entity.Property(e => e.Connombre)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("connombre");
            entity.Property(e => e.DetMe)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("det_me");
            entity.Property(e => e.DetMn)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("det_mn");
            entity.Property(e => e.Direccion)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("direccion");
            entity.Property(e => e.Empcon)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("empcon");
            entity.Property(e => e.Estemp)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estemp");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Indper)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indper");
            entity.Property(e => e.Indret)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("indret");
            entity.Property(e => e.Localidademp)
                .HasMaxLength(40)
                .IsUnicode(false)
                .HasColumnName("localidademp");
            entity.Property(e => e.Nombredbemp)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("nombredbemp");
            entity.Property(e => e.Razemp)
                .HasMaxLength(80)
                .IsUnicode(false)
                .HasColumnName("razemp");
            entity.Property(e => e.Repapematerno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repapematerno");
            entity.Property(e => e.Repapepaterno)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repapepaterno");
            entity.Property(e => e.Repdocumento)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("repdocumento");
            entity.Property(e => e.Repnombre)
                .HasMaxLength(25)
                .IsUnicode(false)
                .HasColumnName("repnombre");
            entity.Property(e => e.Rucemp)
                .HasMaxLength(11)
                .IsUnicode(false)
                .HasColumnName("rucemp");
            entity.Property(e => e.Sincroniza)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sincroniza");
            entity.Property(e => e.Sisban)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sisban");
            entity.Property(e => e.Siscon)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("siscon");
            entity.Property(e => e.Sispla)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("sispla");
            entity.Property(e => e.SmbMe)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("smb_me");
            entity.Property(e => e.SmbMn)
                .HasMaxLength(4)
                .IsUnicode(false)
                .HasColumnName("smb_mn");
            entity.Property(e => e.Usrcre)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrcre");
            entity.Property(e => e.Usrmdf)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasColumnName("usrmdf");
        });

        modelBuilder.Entity<Tgsunat>(entity =>
        {
            entity.HasKey(e => new { e.Codtabla, e.Codsunat }).HasName("PK_tgsunat_tablasunat");

            entity.ToTable("tgsunat", tb => tb.HasComment("Maestro de tablas anexos sunat"));

            entity.Property(e => e.Codtabla)
                .HasMaxLength(3)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codtabla");
            entity.Property(e => e.Codsunat)
                .HasMaxLength(8)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("codsunat");
            entity.Property(e => e.Campo03)
                .HasMaxLength(60)
                .IsUnicode(false)
                .HasColumnName("campo03");
            entity.Property(e => e.Detsunat)
                .HasMaxLength(550)
                .IsUnicode(false)
                .HasColumnName("detsunat");
            entity.Property(e => e.Estsunat)
                .HasMaxLength(1)
                .IsUnicode(false)
                .IsFixedLength()
                .HasColumnName("estsunat");
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

        modelBuilder.Entity<Tgubigeo>(entity =>
        {
            entity.HasKey(e => e.Codubg).HasName("PK_tgubigeo_codubg");

            entity.ToTable("tgubigeo", tb => tb.HasComment("Maestro de ubicacion geografica"));

            entity.Property(e => e.Codubg)
                .HasMaxLength(6)
                .IsUnicode(false)
                .HasColumnName("codubg");
            entity.Property(e => e.Desubg)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("desubg");
            entity.Property(e => e.Fyhcre).HasColumnName("fyhcre");
            entity.Property(e => e.Fyhmdf).HasColumnName("fyhmdf");
            entity.Property(e => e.Nivelubg).HasColumnName("nivelubg");
            entity.Property(e => e.Postalubg)
                .HasMaxLength(5)
                .IsUnicode(false)
                .HasColumnName("postalubg");
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
