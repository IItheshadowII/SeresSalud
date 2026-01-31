using System.Collections.Generic;
using System.ComponentModel;
using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

public class CompanyResolutionForm : Form
{
    private readonly BindingList<OutputRow> _rows;
    private readonly CompanyRepositoryExcel _companyRepository;

    private DataGridView _dgvEmpresas = null!;
    private Button _btnBuscarBase = null!;
    private Button _btnEditarEmpresa = null!;
    private Button _btnCerrar = null!;

    private bool _allowClose;

    public CompanyResolutionForm(IList<OutputRow> rows, CompanyRepositoryExcel companyRepository)
    {
        _rows = new BindingList<OutputRow>(rows);
        _companyRepository = companyRepository;

        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Revisión de datos de empresas";
        Size = new Size(1000, 600);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        FormClosing += CompanyResolutionForm_FormClosing;

        _dgvEmpresas = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            EditMode = DataGridViewEditMode.EditOnEnter,
            ScrollBars = ScrollBars.Both,
            AllowUserToOrderColumns = true
        };

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "CUIT",
            DataPropertyName = nameof(OutputRow.CuitEmpleador),
            Width = 120
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "CIIU",
            DataPropertyName = nameof(OutputRow.CIIU),
            Width = 80
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Empleador",
            DataPropertyName = nameof(OutputRow.Empleador),
            Width = 200
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Calle",
            DataPropertyName = nameof(OutputRow.Calle),
            Width = 180
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Código Postal",
            DataPropertyName = nameof(OutputRow.CodPostal),
            Width = 90
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Localidad",
            DataPropertyName = nameof(OutputRow.Localidad),
            Width = 150
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Provincia",
            DataPropertyName = nameof(OutputRow.Provincia),
            Width = 150
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "ABM loc/prov",
            DataPropertyName = nameof(OutputRow.ABMlocProv),
            Width = 110
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Teléfono",
            DataPropertyName = nameof(OutputRow.Telefono),
            Width = 120
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Fax",
            DataPropertyName = nameof(OutputRow.Fax),
            Width = 120
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Contrato",
            DataPropertyName = nameof(OutputRow.Contrato),
            Width = 90
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Nro Establecimiento",
            DataPropertyName = nameof(OutputRow.NroEstablecimiento),
            Width = 130
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Frecuencia",
            DataPropertyName = nameof(OutputRow.Frecuencia),
            Width = 80
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "CUIL",
            DataPropertyName = nameof(OutputRow.Cuil),
            Width = 120
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Nro Documento",
            DataPropertyName = nameof(OutputRow.NroDocumento),
            Width = 110
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Trabajador",
            DataPropertyName = nameof(OutputRow.TrabajadorApellidoNombre),
            Width = 200
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Riesgo",
            DataPropertyName = nameof(OutputRow.Riesgo),
            Width = 180
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Desc. Riesgo",
            DataPropertyName = nameof(OutputRow.DescripcionRiesgo),
            Width = 180
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "ABM Riesgo",
            DataPropertyName = nameof(OutputRow.ABMRiesgo),
            Width = 110
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Prestación",
            DataPropertyName = nameof(OutputRow.Prestacion),
            Width = 180
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Historia Clínica",
            DataPropertyName = nameof(OutputRow.HistoriaClinica),
            Width = 130
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Mail",
            DataPropertyName = nameof(OutputRow.Mail),
            Width = 180
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Referente",
            DataPropertyName = nameof(OutputRow.Referente),
            Width = 140
        });

        _dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Errores",
            DataPropertyName = nameof(OutputRow.DescripcionError),
            Width = 250,
            ReadOnly = true
        });

        _dgvEmpresas.DataSource = _rows;

        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 70
        };

        _btnBuscarBase = new Button
        {
            Text = "Buscar en Empresas.xlsx",
            Location = new Point(20, 18),
            Size = new Size(180, 35)
        };
        _btnBuscarBase.Click += BtnBuscarBase_Click;

        _btnEditarEmpresa = new Button
        {
            Text = "Editar/Crear empresa",
            Location = new Point(210, 18),
            Size = new Size(160, 35)
        };
        _btnEditarEmpresa.Click += BtnEditarEmpresa_Click;

        _btnCerrar = new Button
        {
            Text = "Aceptar",
            Anchor = AnchorStyles.Top | AnchorStyles.Left,
            Location = new Point(390, 18),
            Size = new Size(150, 35)
        };
        _btnCerrar.Click += (_, _) => { _allowClose = true; DialogResult = DialogResult.OK; Close(); };

        bottomPanel.Controls.Add(_btnBuscarBase);
        bottomPanel.Controls.Add(_btnEditarEmpresa);
        bottomPanel.Controls.Add(_btnCerrar);

        Controls.Add(_dgvEmpresas);
        Controls.Add(bottomPanel);

        AcceptButton = _btnCerrar;
    }

    private void CompanyResolutionForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
        }
    }

    private OutputRow? GetCurrentRow()
    {
        if (_dgvEmpresas.CurrentRow == null)
            return null;

        return _dgvEmpresas.CurrentRow.DataBoundItem as OutputRow;
    }

    private void BtnBuscarBase_Click(object? sender, EventArgs e)
    {
        var row = GetCurrentRow();
        if (row == null)
        {
            MessageBox.Show("Seleccione una fila para buscar en la base.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var companies = new List<CompanyRecord>();

        if (!string.IsNullOrWhiteSpace(row.CuitEmpleador))
        {
            companies = _companyRepository.SearchByCuit(row.CuitEmpleador);
        }
        else if (!string.IsNullOrWhiteSpace(row.Empleador))
        {
            companies = _companyRepository.SearchByName(row.Empleador);
        }

        if (companies.Count == 0)
        {
            MessageBox.Show(
                $"No se encontraron empresas en la base para los datos actuales.\n\n" +
                $"Base: {_companyRepository.FilePath}\n" +
                $"Empresas cargadas: {_companyRepository.CompaniesCount}\n\n" +
                $"Empleador buscado: {row.Empleador}",
                "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (companies.Count == 1)
        {
            ApplyCompanyData(row, companies[0]);
            _dgvEmpresas.Refresh();
            return;
        }

        using var selectDialog = new CompanySelectDialog(companies, _companyRepository);
        if (selectDialog.ShowDialog(this) == DialogResult.OK && selectDialog.SelectedCompany != null)
        {
            ApplyCompanyData(row, selectDialog.SelectedCompany);
            _dgvEmpresas.Refresh();
        }
    }

    private void BtnEditarEmpresa_Click(object? sender, EventArgs e)
    {
        var row = GetCurrentRow();
        if (row == null)
        {
            MessageBox.Show("Seleccione una fila para editar.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var company = new CompanyRecord
        {
            CUIT = row.CuitEmpleador ?? string.Empty,
            CIIU = row.CIIU ?? string.Empty,
            Empleador = row.Empleador ?? string.Empty,
            Calle = row.Calle ?? string.Empty,
            CodPostal = row.CodPostal ?? string.Empty,
            Localidad = row.Localidad ?? string.Empty,
            Provincia = row.Provincia ?? string.Empty,
            Telefono = row.Telefono ?? string.Empty,
            Fax = row.Fax ?? string.Empty,
            Mail = row.Mail ?? string.Empty
        };

        var cuitRequired = string.IsNullOrWhiteSpace(company.CUIT);

        using var editDialog = new CompanyEditDialog(company, cuitRequired);
        if (editDialog.ShowDialog(this) == DialogResult.OK)
        {
            ApplyCompanyData(row, editDialog.Company);
            _companyRepository.SaveCompany(editDialog.Company);
            _dgvEmpresas.Refresh();
        }
    }

    private static void ApplyCompanyData(OutputRow row, CompanyRecord company)
    {
        row.CuitEmpleador = company.CUIT;
        row.CIIU = company.CIIU;
        row.Empleador = company.Empleador;
        row.Calle = company.Calle;
        row.CodPostal = company.CodPostal;
        row.Localidad = company.Localidad;
        row.Provincia = company.Provincia;
        row.Telefono = company.Telefono;
        row.Fax = company.Fax;
        row.Mail = company.Mail;
    }
}
