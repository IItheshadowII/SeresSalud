using ConvertidorDeOrdenes.Desktop.Forms;

namespace ConvertidorDeOrdenes.Desktop;

static class Program
{
    /// <summary>
    /// Punto de entrada principal de la aplicación.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        // Primera ventana: selección de ART (define el flujo completo)
        using var artSelector = new ArtSelectionForm();
        if (artSelector.ShowDialog() != DialogResult.OK)
        {
            return;
        }

        var art = artSelector.ArtSeleccionada;

        // Para futuras ARTs, el flujo puede ser distinto (otros archivos/columnas).
        if (!string.Equals(art, "La Segunda", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(
                $"La ART '{art}' todavía no está implementada.\n\n" +
                "Por ahora, seleccione 'La Segunda'.",
                "No implementado",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        // Mostrar wizard inicial
        using var wizard = new WizardForm();
        
        if (wizard.ShowDialog() == DialogResult.OK)
        {
            // Abrir formulario principal con configuración
            Application.Run(new MainForm(
                wizard.TipoCargaSeleccionado,
                wizard.FrecuenciaSeleccionada,
                wizard.Referente,
                art
            ));
        }
    }
}
