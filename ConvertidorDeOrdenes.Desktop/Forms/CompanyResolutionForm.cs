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
    private Button _btnBuscarCuitOnline = null!;
    private Button _btnEditarEmpresa = null!;
    private Button _btnCerrar = null!;

    public CompanyResolutionForm(IList<OutputRow> rows, CompanyRepositoryExcel companyRepository)
    {
        _rows = new BindingList<OutputRow>(rows);
        _companyRepository = companyRepository;

        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Revisión de datos de empresas";
        Size = new Size(1100, 650);
        MinimumSize = new Size(1100, 650);
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        BackColor = Color.WhiteSmoke;

        var lblTitulo = new Label
        {
            Text = "📋 Revisión de Datos de Empresas y Trabajadores",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Location = new Point(25, 15),
            Size = new Size(1000, 30)
        };

        var lblSubtitulo = new Label
        {
            Text = "Verifique y complete la información antes de exportar. Los campos editables pueden modificarse directamente en la tabla.",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(25, 45),
            Size = new Size(1000, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        this.Controls.Add(lblTitulo);
        this.Controls.Add(lblSubtitulo);

        _dgvEmpresas = new DataGridView
        {
            Location = new Point(25, 75),
            Size = new Size(1040, 470),
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            EditMode = DataGridViewEditMode.EditOnEnter,
            ScrollBars = ScrollBars.Both,
            AllowUserToOrderColumns = true,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            ColumnHeadersHeight = 35,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(70, 130, 180),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Padding = new Padding(5),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                WrapMode = DataGridViewTriState.False
            },
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(245, 245, 245)
            }
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
            Height = 80,
            BackColor = Color.WhiteSmoke
        };

        // Agregar panel al formulario primero para que tenga el ancho correcto
        Controls.Add(_dgvEmpresas);
        Controls.Add(bottomPanel);

        _btnBuscarBase = new Button
        {
            Text = "Buscar en Empresas.xlsx",
            Location = new Point(25, 20),
            Size = new Size(200, 40),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnBuscarBase.FlatAppearance.BorderSize = 0;
        _btnBuscarBase.Click += BtnBuscarBase_Click;

        _btnBuscarCuitOnline = new Button
        {
            Text = "Buscar CUIT online",
            Location = new Point(235, 20),
            Size = new Size(180, 40),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(100, 100, 160),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnBuscarCuitOnline.FlatAppearance.BorderSize = 0;
        _btnBuscarCuitOnline.Click += BtnBuscarCuitOnline_Click;

        _btnEditarEmpresa = new Button
        {
            Text = "Editar/Crear empresa",
            Location = new Point(425, 20),
            Size = new Size(180, 40),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnEditarEmpresa.FlatAppearance.BorderSize = 0;
        _btnEditarEmpresa.Click += BtnEditarEmpresa_Click;

        _btnCerrar = new Button
        {
            Text = "Aceptar",
            Anchor = AnchorStyles.Right,
            Location = new Point(bottomPanel.ClientSize.Width - 225, 20),
            Size = new Size(200, 40),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnCerrar.FlatAppearance.BorderSize = 0;
        _btnCerrar.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };

        bottomPanel.Controls.Add(_btnBuscarBase);
        bottomPanel.Controls.Add(_btnBuscarCuitOnline);
        bottomPanel.Controls.Add(_btnEditarEmpresa);
        bottomPanel.Controls.Add(_btnCerrar);

        AcceptButton = _btnCerrar;
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

    private void BtnBuscarCuitOnline_Click(object? sender, EventArgs e)
    {
        var row = GetCurrentRow();
        if (row == null)
        {
            MessageBox.Show("Seleccione una fila para buscar el CUIT online.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (string.IsNullOrWhiteSpace(row.Contrato))
        {
            MessageBox.Show("La fila seleccionada no tiene N° de contrato. Complete el contrato antes de buscar el CUIT online.",
                "Contrato requerido",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        using var dlg = new CuitOnlineLookupForm(row.Contrato);
        if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.FoundCuit))
        {
            row.CuitEmpleador = dlg.FoundCuit;
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
