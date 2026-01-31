using System.ComponentModel;
using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Formulario para administrar la base de Empresas.xlsx (listar, agregar, editar, eliminar).
/// </summary>
public sealed class CompanyListForm : Form
{
    private readonly CompanyRepositoryExcel _repository;
    private List<CompanyRecord> _allCompanies = new();
    private BindingList<CompanyRecord> _binding = null!;

    private TextBox _txtBuscar = null!;
    private DataGridView _dgv = null!;
    private Button _btnAgregar = null!;
    private Button _btnEditar = null!;
    private Button _btnEliminar = null!;
    private Button _btnCerrar = null!;

    public CompanyListForm(CompanyRepositoryExcel repository)
    {
        _repository = repository;
        InitializeComponent();
        LoadCompanies();
    }

    private void InitializeComponent()
    {
        Text = "Empresas (Empresas.xlsx)";
        Size = new Size(900, 500);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _txtBuscar = new TextBox
        {
            Location = new Point(20, 15),
            Size = new Size(500, 23),
            PlaceholderText = "Buscar por CUIT, nombre o localidad..."
        };
        _txtBuscar.TextChanged += (_, _) => ApplyFilter();

        _dgv = new DataGridView
        {
            Location = new Point(20, 45),
            Size = new Size(840, 360),
            AutoGenerateColumns = true,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false
        };

        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60
        };

        _btnAgregar = new Button
        {
            Text = "Agregar",
            Location = new Point(20, 15),
            Size = new Size(100, 30)
        };
        _btnAgregar.Click += BtnAgregar_Click;

        _btnEditar = new Button
        {
            Text = "Editar",
            Location = new Point(130, 15),
            Size = new Size(100, 30)
        };
        _btnEditar.Click += BtnEditar_Click;

        _btnEliminar = new Button
        {
            Text = "Eliminar",
            Location = new Point(240, 15),
            Size = new Size(100, 30)
        };
        _btnEliminar.Click += BtnEliminar_Click;

        _btnCerrar = new Button
        {
            Text = "Cerrar",
            Location = new Point(760, 15),
            Size = new Size(100, 30)
        };
        _btnCerrar.Click += (_, _) => Close();

        bottomPanel.Controls.Add(_btnAgregar);
        bottomPanel.Controls.Add(_btnEditar);
        bottomPanel.Controls.Add(_btnEliminar);
        bottomPanel.Controls.Add(_btnCerrar);
        Controls.Add(_txtBuscar);
        Controls.Add(_dgv);
        Controls.Add(bottomPanel);
    }

    private void LoadCompanies()
    {
        _allCompanies = _repository.GetAll();
        ApplyFilter();
    }

    private CompanyRecord? GetSelectedCompany()
    {
        if (_dgv.CurrentRow?.DataBoundItem is CompanyRecord company)
            return company;

        return null;
    }

    private void BtnAgregar_Click(object? sender, EventArgs e)
    {
        var newCompany = new CompanyRecord();
        using var dlg = new CompanyEditDialog(newCompany, cuitRequired: true);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _repository.SaveCompany(dlg.Company);
            LoadCompanies();
        }
    }

    private void BtnEditar_Click(object? sender, EventArgs e)
    {
        var selected = GetSelectedCompany();
        if (selected == null)
        {
            MessageBox.Show("Seleccione una empresa para editar.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Trabajar sobre una copia para no modificar la lista hasta confirmar.
        var clone = new CompanyRecord
        {
            RowIndex = selected.RowIndex,
            CUIT = selected.CUIT,
            CIIU = selected.CIIU,
            Empleador = selected.Empleador,
            Calle = selected.Calle,
            CodPostal = selected.CodPostal,
            Localidad = selected.Localidad,
            Provincia = selected.Provincia,
            Telefono = selected.Telefono,
            Fax = selected.Fax,
            Mail = selected.Mail
        };

        using var dlg = new CompanyEditDialog(clone, cuitRequired: !string.IsNullOrWhiteSpace(clone.CUIT));
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _repository.SaveCompany(dlg.Company);
            LoadCompanies();
        }
    }

    private void BtnEliminar_Click(object? sender, EventArgs e)
    {
        var selected = GetSelectedCompany();
        if (selected == null)
        {
            MessageBox.Show("Seleccione una empresa para eliminar.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var confirm = MessageBox.Show(
            $"¿Confirma que desea eliminar la empresa:\n\n{selected.CUIT} - {selected.Empleador}?\n\nSe realizará un backup previo de Empresas.xlsx.",
            "Confirmar eliminación",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
            return;

        try
        {
            _repository.DeleteCompany(selected);
            LoadCompanies();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error eliminando empresa: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ApplyFilter()
    {
        var term = (_txtBuscar.Text ?? string.Empty).Trim().ToUpperInvariant();

        List<CompanyRecord> filtered;

        if (string.IsNullOrWhiteSpace(term))
        {
            filtered = _allCompanies;
        }
        else
        {
            filtered = _allCompanies
                .Where(c =>
                {
                    var composite = $"{c.CUIT} {c.Empleador} {c.Localidad} {c.Provincia}".ToUpperInvariant();
                    return composite.Contains(term);
                })
                .ToList();
        }

        _binding = new BindingList<CompanyRecord>(filtered);
        _dgv.DataSource = _binding;
        if (_dgv.Columns.Contains(nameof(CompanyRecord.RowIndex)))
        {
            _dgv.Columns[nameof(CompanyRecord.RowIndex)].Visible = false;
        }
    }
}
