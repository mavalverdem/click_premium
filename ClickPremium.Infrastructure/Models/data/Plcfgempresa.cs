using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Configuracion de parametros de empresa
/// </summary>
public partial class Plcfgempresa
{
    public string Pdoano { get; set; } = null!;

    public string? Codvia { get; set; }

    public string? Direccionvia { get; set; }

    public string? Numerodir { get; set; }

    public string? Codzona { get; set; }

    public string? Direccionzona { get; set; }

    public string? Ubigeodir { get; set; }

    public string? Regpatronal { get; set; }

    public string? Girocomercial { get; set; }

    public string? Telefono { get; set; }

    public string? Email { get; set; }

    public string? Repapepaterno { get; set; }

    public string? Repapematerno { get; set; }

    public string? Repnombres { get; set; }

    public string? Repcargo { get; set; }

    public string? Repcoddci { get; set; }

    public string? Repnumdocu { get; set; }

    public string? Gerapepaterno { get; set; }

    public string? Gerapematerno { get; set; }

    public string? Gernombres { get; set; }

    public string? Gercargo { get; set; }

    public string? Gercoddci { get; set; }

    public string? Gernumdocu { get; set; }

    public string? Psnapepaterno { get; set; }

    public string? Psnapematerno { get; set; }

    public string? Psnnombres { get; set; }

    public string? Psntelefono { get; set; }

    public string? Codcco { get; set; }

    public short ServerEnvio { get; set; }

    public string? UsuarioEnvio { get; set; }

    public string? PasswordEnvio { get; set; }

    public string? CorreoEnvio { get; set; }

    public short PuertoEnvio { get; set; }

    public string? Codcpcrem { get; set; }

    public string? Codtbluit { get; set; }

    public string? Codcpc5ta { get; set; }

    public string? Codcpc5taIng { get; set; }

    public string Repimpbol { get; set; } = null!;

    public string Dirimpbol { get; set; } = null!;

    public string? ContratoDot { get; set; }

    public string? ContratoDoc { get; set; }

    public string? Rembasica { get; set; }

    public string? Rempromedio { get; set; }

    public string? Remempordin { get; set; }

    public string? Remempextra { get; set; }

    public string? Rempendiente { get; set; }

    public string? Gratipendiente { get; set; }

    public string? Remanterior { get; set; }

    public string? Remganada { get; set; }

    public string? Codtblretener { get; set; }

    public string? Codtblpendiente { get; set; }

    public string? Codtbldividir { get; set; }

    public string Gratixasis { get; set; } = null!;

    public string Gratiliqxdias { get; set; } = null!;

    public string? Remxutiejer1 { get; set; }

    public string? Remxutiejer2 { get; set; }

    public string? Remxutiejer3 { get; set; }

    public string? Remxutiejer4 { get; set; }

    public decimal RentaxejerMn { get; set; }

    public decimal RentaxejerMe { get; set; }

    public decimal Porcepartici { get; set; }

    public short Nivelcencosto { get; set; }

    public string LiqprnRazonemp { get; set; } = null!;

    public string LiqprnLogoemp { get; set; } = null!;

    public byte[]? Logo { get; set; }

    public byte[]? Firma { get; set; }

    public byte[]? Firmanexo { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Pltipovium? CodviaNavigation { get; set; }

    public virtual Pltipozona? CodzonaNavigation { get; set; }

    public virtual Pldocidentidad? GercoddciNavigation { get; set; }

    public virtual Pldocidentidad? RepcoddciNavigation { get; set; }
}
