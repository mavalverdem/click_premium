using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Familiares Trabajador
/// </summary>
public partial class Plfamiliare
{
    public string Codcls { get; set; } = null!;

    public string Codpsn { get; set; } = null!;

    public short Orden { get; set; }

    public string? Apepaterno { get; set; }

    public string? Apematerno { get; set; }

    public string? Nombres { get; set; }

    public DateTime? Fecnacimiento { get; set; }

    public string Sexofam { get; set; } = null!;

    public string? Coddci { get; set; }

    public string? Numdociden { get; set; }

    public string? Vinculo { get; set; }

    public string? Cartamed { get; set; }

    public string Domicilio { get; set; } = null!;

    public string? Codvia { get; set; }

    public string? Nomviadom { get; set; }

    public string? Numerdom { get; set; }

    public string? Intedom { get; set; }

    public string? Codzona { get; set; }

    public string? Nomzonadom { get; set; }

    public string? Refedom { get; set; }

    public string? Ubigeodom { get; set; }

    public string Incapacidad { get; set; } = null!;

    public string? Certificadomed { get; set; }

    public string? Motivoina { get; set; }

    public string? Tipdocpaternidad { get; set; }

    public string? Acrepaternidad { get; set; }

    public DateTime? Fecalta { get; set; }

    public DateTime? Fecbaja { get; set; }

    public string Estadofam { get; set; } = null!;

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }
}
