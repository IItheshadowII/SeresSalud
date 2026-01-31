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
        Text = "Administración de Empresas";
        Size = new Size(980, 580);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.WhiteSmoke;

        var lblTitulo = new Label
        {
            Text = "🏢 Administración de Empresas",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Location = new Point(25, 15),
            Size = new Size(500, 30)
        };

        var lblSubtitulo = new Label
        {
            Text = "Gestione la base de datos de empresas (Empresas.xlsx). Agregue, edite o elimine registros.",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(25, 45),
            Size = new Size(900, 18)
        };

        _txtBuscar = new TextBox
        {
            Location = new Point(25, 70),
            Size = new Size(600, 28),
            Font = new Font("Segoe UI", 10),
            PlaceholderText = "🔍 Buscar por CUIT, nombre o localidad..."
        };
        _txtBuscar.TextChanged += (_, _) => ApplyFilter();

        _dgv = new DataGridView
        {
            Location = new Point(25, 110),
            Size = new Size(920, 365),
            AutoGenerateColumns = true,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeight = 35,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(70, 130, 180),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                WrapMode = DataGridViewTriState.False
            }
        };

        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 70,
            BackColor = Color.WhiteSmoke
        };

        _btnAgregar = new Button
        {
            Text = "Agregar",
            Location = new Point(25, 15),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnAgregar.FlatAppearance.BorderSize = 0;
        _btnAgregar.Click += BtnAgregar_Click;

        _btnEditar = new Button
        {
            Text = "Editar",
            Location = new Point(145, 15),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnEditar.FlatAppearance.BorderSize = 0;
        _btnEditar.Click += BtnEditar_Click;

        _btnEliminar = new Button
        {
            Text = "Eliminar",
            Location = new Point(265, 15),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(200, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnEliminar.FlatAppearance.BorderSize = 0;
        _btnEliminar.Click += BtnEliminar_Click;

        _btnCerrar = new Button
        {
            Text = "Cerrar",
            Location = new Point(835, 15),
            Size = new Size(110, 38),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnCerrar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        _btnCerrar.Click += (_, _) => Close();

        bottomPanel.Controls.Add(_btnAgregar);
        bottomPanel.Controls.Add(_btnEditar);
        bottomPanel.Controls.Add(_btnEliminar);
        bottomPanel.Controls.Add(_btnCerrar);
        Controls.Add(lblTitulo);
        Controls.Add(lblSubtitulo);
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
