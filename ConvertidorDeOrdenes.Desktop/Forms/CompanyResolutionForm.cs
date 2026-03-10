using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;
using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

public class CompanyResolutionForm : Form
{
    private readonly BindingList<OutputRow> _rows;
    private readonly CompanyRepositoryExcel _companyRepository;
    private readonly CompanyResolutionService _resolutionService;

    private DataGridView _dgvEmpresas = null!;
    private Button _btnBuscarBase = null!;
    private Button _btnBuscarCuitOnline = null!;
    private Button _btnEditarEmpresa = null!;
    private Button _btnRefrescar = null!;
    private Button _btnCerrar = null!;

    public CompanyResolutionForm(IList<OutputRow> rows, CompanyRepositoryExcel companyRepository)
    {
        _rows = new BindingList<OutputRow>(rows);
        _companyRepository = companyRepository;
        _resolutionService = new CompanyResolutionService(companyRepository);

        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Revisión de datos de empresas";
        Icon = AppIcon.TryGet();
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
            Width = 90,
            ReadOnly = true
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

        _btnRefrescar = new Button
        {
            Text = "Refrescar datos",
            Location = new Point(615, 20),
            Size = new Size(180, 40),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(255, 215, 115),
            ForeColor = Color.FromArgb(80, 60, 0),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnRefrescar.FlatAppearance.BorderSize = 0;
        _btnRefrescar.Click += BtnRefrescar_Click;

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
        _btnCerrar.Click += BtnAceptar_Click;

        bottomPanel.Controls.Add(_btnBuscarBase);
        bottomPanel.Controls.Add(_btnBuscarCuitOnline);
        bottomPanel.Controls.Add(_btnEditarEmpresa);
        bottomPanel.Controls.Add(_btnRefrescar);
        bottomPanel.Controls.Add(_btnCerrar);

        AcceptButton = _btnCerrar;
    }

    private OutputRow? GetCurrentRow()
    {
        if (_dgvEmpresas.CurrentRow == null)
            return null;

        return _dgvEmpresas.CurrentRow.DataBoundItem as OutputRow;
    }

    private void BtnAceptar_Click(object? sender, EventArgs e)
    {
        try
        {
            PersistResolvedCompanies();
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error guardando empresas detectadas: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
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
            var company = companies[0];
            ApplyCompanyData(row, company);
            ApplyCompanyToSimilarRows(row, company);
            _dgvEmpresas.Refresh();
            return;
        }

        using var selectDialog = new CompanySelectDialog(companies, _companyRepository);
        if (selectDialog.ShowDialog(this) == DialogResult.OK && selectDialog.SelectedCompany != null)
        {
            var selectedCompany = selectDialog.SelectedCompany;
            ApplyCompanyData(row, selectedCompany);
            ApplyCompanyToSimilarRows(row, selectedCompany);
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
            RowIndex = row.ResolvedCompanyRowIndex,
            CUIT = row.CuitEmpleador ?? string.Empty,
            CIIU = row.CIIU ?? string.Empty,
            Empleador = row.Empleador ?? string.Empty,
            Calle = row.Calle ?? string.Empty,
            CodPostal = string.IsNullOrWhiteSpace(row.ResolvedCompanyCodPostal) ? row.CodPostal ?? string.Empty : row.ResolvedCompanyCodPostal,
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
            _resolutionService.SaveWithResolution(this, editDialog.Company);
            _dgvEmpresas.Refresh();
        }
    }

    private async void BtnRefrescar_Click(object? sender, EventArgs e)
    {
        if (_rows.Count == 0)
            return;

        using var progress = new ProgressDialog("Actualizando", "Reaplicando datos de Empresas.xlsx...");
        progress.SetIndeterminate(false);
        progress.Show(this);
        await Task.Yield();

        progress.SetStatus($"Reaplicando datos de Empresas.xlsx... 0 / {_rows.Count} filas");

        var resolveProgress = new Progress<(int processed, int total)>(info =>
        {
            var (processed, total) = info;
            var percent = total > 0 ? (int)Math.Round(processed * 100.0 / total) : 0;
            progress.SetProgress(percent);
            progress.SetStatus($"Reaplicando datos de Empresas.xlsx... {processed} / {total} filas");
        });

        await Task.Run(() => AutoResolveCompaniesOnRows(resolveProgress));

        _dgvEmpresas.Refresh();
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
            var foundCuit = dlg.FoundCuit!;
            row.CuitEmpleador = foundCuit;

            // Si el CUIT extraído ya existe en Empresas.xlsx, usamos esa empresa como
            // fuente de verdad y la aplicamos al resto de filas similares.
            var companies = _companyRepository.SearchByCuit(foundCuit);
            var resolvedByAddress = ResolveCompanyByCuitAndAddress(companies, row);
            if (resolvedByAddress != null)
            {
                ApplyCompanyData(row, resolvedByAddress);
                ApplyCompanyToSimilarRows(row, resolvedByAddress);
            }
            else
            {
                // Si todavía no está en Empresas.xlsx, al menos propagamos el CUIT
                // al resto de filas con el mismo empleador / CUIT vacío.
                ApplyCuitToSimilarRows(row, foundCuit);
            }

            _dgvEmpresas.Refresh();
        }
    }

    private void AutoResolveCompaniesOnRows(IProgress<(int processed, int total)>? progress = null)
    {
        var total = _rows.Count;
        for (int i = 0; i < total; i++)
        {
            var row = _rows[i];
            ExtractCodPostalFromLocalidad(row);
            if (string.IsNullOrWhiteSpace(row.ResolvedCompanyCodPostal) && !string.IsNullOrWhiteSpace(row.CodPostal))
                row.ResolvedCompanyCodPostal = row.CodPostal;

            // Limpiamos la localidad pero preservamos el CP extraído desde la orden.
            StripCodPostalFromLocalidad(row);

            if (!string.IsNullOrWhiteSpace(row.CuitEmpleador))
            {
                var companies = _companyRepository.SearchByCuit(row.CuitEmpleador);

                var company = ResolveCompanyByCuitAndAddress(companies, row);
                if (company != null)
                {
                    ApplyCompanyDataPreservingSede(row, company);
                }
            }
            else if (!string.IsNullOrWhiteSpace(row.Empleador))
            {
                var companies = _companyRepository.SearchByName(row.Empleador);

                if (companies.Count == 1)
                {
                    var company = companies[0];
                    ApplyCompanyDataPreservingSede(row, company);
                }
            }

            if (progress != null && total > 0 && (i % 250 == 0 || i == total - 1))
            {
                progress.Report((i + 1, total));
            }
        }
    }

    private static string GetCuitDigits(string? cuit)
    {
        if (string.IsNullOrWhiteSpace(cuit))
            return string.Empty;

        return new string(cuit.Where(char.IsDigit).ToArray());
    }

    private static string NormalizeName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return value.Trim().ToUpperInvariant();
    }

    private static bool TextEquals(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return false;

        return left.Trim().Equals(right.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAddressCompatibleForPropagation(OutputRow sourceRow, OutputRow targetRow)
    {
        // Si el destino no tiene datos de dirección/sede, se permite completar.
        var targetHasAddress =
            !string.IsNullOrWhiteSpace(targetRow.Calle) ||
            !string.IsNullOrWhiteSpace(targetRow.Localidad) ||
            !string.IsNullOrWhiteSpace(targetRow.Provincia);

        if (!targetHasAddress)
            return true;

        // Si ambos campos existen y difieren, consideramos que son sedes distintas.
        if (!string.IsNullOrWhiteSpace(sourceRow.Provincia) && !string.IsNullOrWhiteSpace(targetRow.Provincia) &&
            !TextEquals(sourceRow.Provincia, targetRow.Provincia))
            return false;

        if (!string.IsNullOrWhiteSpace(sourceRow.Localidad) && !string.IsNullOrWhiteSpace(targetRow.Localidad) &&
            !TextEquals(sourceRow.Localidad, targetRow.Localidad))
            return false;

        if (!string.IsNullOrWhiteSpace(sourceRow.Calle) && !string.IsNullOrWhiteSpace(targetRow.Calle) &&
            !TextEquals(sourceRow.Calle, targetRow.Calle))
            return false;

        return true;
    }

    private static CompanyRecord? ResolveCompanyByCuitAndAddress(IReadOnlyList<CompanyRecord> companies, OutputRow row)
    {
        if (companies.Count == 1)
            return companies[0];

        if (companies.Count <= 1)
            return null;

        var byProvincia = companies
            .Where(c => TextEquals(c.Provincia, row.Provincia))
            .ToList();

        if (byProvincia.Count == 1)
            return byProvincia[0];

        var provinceScope = byProvincia.Count > 1 ? byProvincia : companies;

        var byLocalidad = provinceScope
            .Where(c => TextEquals(c.Localidad, row.Localidad))
            .ToList();

        if (byLocalidad.Count == 1)
            return byLocalidad[0];

        var localidadScope = byLocalidad.Count > 1 ? byLocalidad : provinceScope;

        var byCalle = localidadScope
            .Where(c => TextEquals(c.Calle, row.Calle))
            .ToList();

        return byCalle.Count == 1 ? byCalle[0] : null;
    }

    /// <summary>
    /// Aplica la empresa elegida también al resto de filas similares del archivo actual
    /// (mismo CUIT; o mismo empleador cuando la fila no tiene CUIT).
    /// </summary>
    private void ApplyCompanyToSimilarRows(OutputRow sourceRow, CompanyRecord company)
    {
        var companyCuitDigits = GetCuitDigits(company.CUIT);
        var companyNameKey = NormalizeName(company.Empleador);

        if (string.IsNullOrEmpty(companyCuitDigits) && string.IsNullOrEmpty(companyNameKey))
            return;

        foreach (var row in _rows)
        {
            if (ReferenceEquals(row, sourceRow))
                continue;

            var rowCuitDigits = GetCuitDigits(row.CuitEmpleador);

            bool shouldApply = false;

            // Si ya tiene CUIT y coincide exactamente, usamos esa empresa.
            if (!string.IsNullOrEmpty(companyCuitDigits) && !string.IsNullOrEmpty(rowCuitDigits) &&
                rowCuitDigits == companyCuitDigits)
            {
                shouldApply = IsAddressCompatibleForPropagation(sourceRow, row);
            }
            // Si la fila no tiene CUIT pero el nombre de empleador coincide, también aplicamos.
            else if (string.IsNullOrEmpty(rowCuitDigits) && !string.IsNullOrEmpty(companyNameKey) &&
                     NormalizeName(row.Empleador) == companyNameKey)
            {
                shouldApply = IsAddressCompatibleForPropagation(sourceRow, row);
            }

            if (!shouldApply)
                continue;

            ApplyCompanyData(row, company);
        }
    }

    /// <summary>
    /// Cuando solo tenemos el CUIT (por ejemplo, extraído online), lo propagamos a
    /// filas similares: mismo CUIT o mismo empleador sin CUIT.
    /// </summary>
    private void ApplyCuitToSimilarRows(OutputRow sourceRow, string formattedCuit)
    {
        var cuitDigits = GetCuitDigits(formattedCuit);
        if (string.IsNullOrEmpty(cuitDigits))
            return;

        var sourceNameKey = NormalizeName(sourceRow.Empleador);

        foreach (var row in _rows)
        {
            if (ReferenceEquals(row, sourceRow))
                continue;

            var rowCuitDigits = GetCuitDigits(row.CuitEmpleador);
            var rowNameKey = NormalizeName(row.Empleador);

            bool shouldApply = false;

            if (!string.IsNullOrEmpty(rowCuitDigits) && rowCuitDigits == cuitDigits)
            {
                shouldApply = true;
            }
            else if (string.IsNullOrEmpty(rowCuitDigits) && !string.IsNullOrEmpty(sourceNameKey) &&
                     rowNameKey == sourceNameKey)
            {
                shouldApply = true;
            }

            if (!shouldApply)
                continue;

            row.CuitEmpleador = formattedCuit;
        }
    }

    private static void StripCodPostalFromLocalidad(OutputRow row)
    {
        if (string.IsNullOrWhiteSpace(row.Localidad))
            return;

        var match = Regex.Match(row.Localidad, "^\\((\\d{3,5})\\)\\s*(.+)$");
        var localidad = match.Success ? match.Groups[2].Value.Trim() : row.Localidad;
        row.Localidad = NormalizeLocalidadForAnalysis(localidad);
    }

    private static void ExtractCodPostalFromLocalidad(OutputRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.CodPostal))
            return;

        if (string.IsNullOrWhiteSpace(row.Localidad))
            return;

        var match = Regex.Match(row.Localidad, "^\\((\\d{3,5})\\)\\s*(.+)$");
        if (!match.Success)
            return;

        row.CodPostal = match.Groups[1].Value.Trim();
        row.Localidad = match.Groups[2].Value.Trim();
    }

    private void PersistResolvedCompanies()
    {
        var uniqueCompanies = _rows
            .Where(HasPersistableCompanyData)
            .Select(CreateCompanyRecordFromRow)
            .GroupBy(BuildCompanyIdentityKey, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        foreach (var company in uniqueCompanies)
        {
            PersistCompanyRecord(company);
        }
    }

    private void PersistCompanyRecord(CompanyRecord company)
    {
        if (string.IsNullOrWhiteSpace(company.CUIT) || string.IsNullOrWhiteSpace(company.Empleador))
            return;

        if (company.RowIndex <= 0 && !CompanySedeUtils.HasComparableSedeData(company))
            return;

        _companyRepository.SaveCompany(company, forceNew: false);
    }

    private static bool HasPersistableCompanyData(OutputRow row)
    {
        return !string.IsNullOrWhiteSpace(row.CuitEmpleador) &&
               !string.IsNullOrWhiteSpace(row.Empleador) &&
               (row.ResolvedCompanyRowIndex > 0 ||
                CompanySedeUtils.HasComparableSedeData(row.Calle, row.Localidad, row.Provincia));
    }

    private static CompanyRecord CreateCompanyRecordFromRow(OutputRow row)
    {
        return new CompanyRecord
        {
            RowIndex = row.ResolvedCompanyRowIndex,
            CUIT = row.CuitEmpleador,
            CIIU = row.CIIU,
            Empleador = row.Empleador,
            Calle = row.Calle,
            CodPostal = string.IsNullOrWhiteSpace(row.ResolvedCompanyCodPostal) ? row.CodPostal : row.ResolvedCompanyCodPostal,
            Localidad = row.Localidad,
            Provincia = row.Provincia,
            Telefono = row.Telefono,
            Fax = row.Fax,
            Mail = row.Mail
        };
    }

    private static string BuildCompanyIdentityKey(CompanyRecord company)
    {
        if (company.RowIndex > 0)
            return $"ROW:{company.RowIndex}";

        var cuit = GetCuitDigits(company.CUIT);
        var sede = CompanySedeUtils.ComputeSedeKey(company);
        return $"{cuit}|{sede}";
    }

    private static string NormalizeLocalidadForAnalysis(string? localidad)
    {
        if (string.IsNullOrWhiteSpace(localidad))
            return string.Empty;

        var normalized = localidad.Trim();
        normalized = Regex.Replace(normalized, @"^\(\d+\)\s*", string.Empty);
        normalized = Regex.Replace(normalized, @"[-\s]+(B\s*A|BUENOS\s+AIRES)\s*$", string.Empty, RegexOptions.IgnoreCase);
        return normalized.Trim();
    }

    private static void ApplyCompanyData(OutputRow row, CompanyRecord company)
    {
        row.ResolvedCompanyRowIndex = company.RowIndex;
        row.ResolvedCompanyCodPostal = company.CodPostal;
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

    private static void ApplyCompanyDataPreservingSede(OutputRow row, CompanyRecord company)
    {
        row.ResolvedCompanyRowIndex = company.RowIndex;
        row.ResolvedCompanyCodPostal = company.CodPostal;
        row.CuitEmpleador = company.CUIT;
        row.CIIU = company.CIIU;
        row.Empleador = company.Empleador;

        // En refrescos automáticos no pisamos la sede ya informada en la fila
        // para evitar unificar erróneamente distintas plantas del mismo CUIT.
        if (string.IsNullOrWhiteSpace(row.Calle) && !string.IsNullOrWhiteSpace(company.Calle))
            row.Calle = company.Calle;

        if (string.IsNullOrWhiteSpace(row.CodPostal) && !string.IsNullOrWhiteSpace(company.CodPostal))
            row.CodPostal = company.CodPostal;

        if (string.IsNullOrWhiteSpace(row.Localidad) && !string.IsNullOrWhiteSpace(company.Localidad))
            row.Localidad = company.Localidad;

        if (string.IsNullOrWhiteSpace(row.Provincia) && !string.IsNullOrWhiteSpace(company.Provincia))
            row.Provincia = company.Provincia;

        if (string.IsNullOrWhiteSpace(row.Telefono) && !string.IsNullOrWhiteSpace(company.Telefono))
            row.Telefono = company.Telefono;

        if (string.IsNullOrWhiteSpace(row.Fax) && !string.IsNullOrWhiteSpace(company.Fax))
            row.Fax = company.Fax;

        if (string.IsNullOrWhiteSpace(row.Mail) && !string.IsNullOrWhiteSpace(company.Mail))
            row.Mail = company.Mail;
    }
}
