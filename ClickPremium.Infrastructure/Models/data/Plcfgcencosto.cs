using System;
using System.Collections.Generic;

namespace ClickPremium.Infrastructure.Models.data;

/// <summary>
/// Maestro Informacion del Negocio centro costo
/// </summary>
public partial class Plcfgcencosto
{
    public string Codcco { get; set; } = null!;

    public string? Lineanegocio { get; set; }

    public string? Segmentonego { get; set; }

    public string? Clientenego { get; set; }

    public string Usrcre { get; set; } = null!;

    public DateTime Fyhcre { get; set; }

    public string? Usrmdf { get; set; }

    public DateTime? Fyhmdf { get; set; }

    public virtual Cocco CodccoNavigation { get; set; } = null!;
}
