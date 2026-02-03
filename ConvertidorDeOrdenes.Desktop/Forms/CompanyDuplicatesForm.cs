using System.ComponentModel;
using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

/// <summary>
/// Ventana para revisar y depurar duplicados en Empresas.xlsx
/// (mismo CUIT y misma sede: Calle + Localidad + Provincia).
/// Permite unificar (eliminando duplicados) y eliminar registros seleccionados.
/// </summary>
public sealed class CompanyDuplicatesForm : Form
{
    private readonly CompanyRepositoryExcel _repository;

    private readonly BindingList<DuplicateRowView> _binding = new();

    private DataGridView _dgv = null!;
    private Label _lblResumen = null!;
    private Button _btnUnificar = null!;
    private Button _btnEliminar = null!;
    private Button _btnCerrar = null!;

    public CompanyDuplicatesForm(CompanyRepositoryExcel repository)
    {
        _repository = repository;
        InitializeComponent();
        LoadDuplicates();
    }

    private void InitializeComponent()
    {
        Text = "Verificar duplicados de Empresas";
        Icon = AppIcon.TryGet();
        Size = new Size(1100, 600);
        MinimumSize = new Size(1100, 600);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.WhiteSmoke;

        var lblTitulo = new Label
        {
            Text = "🔎 Verificación de Duplicados (Empresas.xlsx)",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50),
            Location = new Point(25, 15),
            Size = new Size(800, 28)
        };

        var lblSubtitulo = new Label
        {
            Text = "Se consideran duplicados los registros con mismo CUIT y misma sede (Calle + Localidad + Provincia).",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(25, 45),
            Size = new Size(1000, 18)
        };

        _lblResumen = new Label
        {
            Text = "Analizando...",
            Font = new Font("Segoe UI", 9, FontStyle.Italic),
            ForeColor = Color.FromArgb(90, 90, 90),
            Location = new Point(25, 68),
            Size = new Size(1000, 18)
        };

        _dgv = new DataGridView
        {
            Location = new Point(25, 95),
            Size = new Size(1040, 390),
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
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
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                WrapMode = DataGridViewTriState.False
            },
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(245, 245, 245)
            }
        };

        // Columnas
        var colSel = new DataGridViewCheckBoxColumn
        {
            HeaderText = "✔",
            DataPropertyName = nameof(DuplicateRowView.Selected),
            Width = 35
        };
        _dgv.Columns.Add(colSel);

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Grupo",
            DataPropertyName = nameof(DuplicateRowView.GroupId),
            Width = 60,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "CUIT",
            DataPropertyName = nameof(DuplicateRowView.CUIT),
            Width = 120,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Empleador",
            DataPropertyName = nameof(DuplicateRowView.Empleador),
            Width = 220,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Calle",
            DataPropertyName = nameof(DuplicateRowView.Calle),
            Width = 190,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Localidad",
            DataPropertyName = nameof(DuplicateRowView.Localidad),
            Width = 150,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Provincia",
            DataPropertyName = nameof(DuplicateRowView.Provincia),
            Width = 150,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "En grupo",
            DataPropertyName = nameof(DuplicateRowView.GroupCount),
            Width = 70,
            ReadOnly = true
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "RowIndex",
            DataPropertyName = nameof(DuplicateRowView.RowIndex),
            Visible = false
        });

        _dgv.DataSource = _binding;

        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 70,
            BackColor = Color.WhiteSmoke
        };

        _btnUnificar = new Button
        {
            Text = "Unificar duplicados (eliminar seleccionados)",
            Location = new Point(25, 15),
            Size = new Size(280, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnUnificar.FlatAppearance.BorderSize = 0;
        _btnUnificar.Click += async (_, _) => await DeleteSelected(confirmText: "¿Desea eliminar los registros seleccionados para unificar los duplicados?\n\nSe mantendrá al menos un registro por grupo.");

        _btnEliminar = new Button
        {
            Text = "Eliminar seleccionados",
            Location = new Point(315, 15),
            Size = new Size(190, 38),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(200, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _btnEliminar.FlatAppearance.BorderSize = 0;
        _btnEliminar.Click += async (_, _) => await DeleteSelected(confirmText: "¿Confirma que desea eliminar definitivamente los registros seleccionados?\n\nSe realizará un backup previo de Empresas.xlsx.");

        _btnCerrar = new Button
        {
            Text = "Cerrar",
            Location = new Point(900, 15),
            Size = new Size(140, 38),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(100, 100, 100),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _btnCerrar.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        _btnCerrar.Click += (_, _) => Close();

        bottomPanel.Controls.Add(_btnUnificar);
        bottomPanel.Controls.Add(_btnEliminar);
        bottomPanel.Controls.Add(_btnCerrar);

        Controls.Add(lblTitulo);
        Controls.Add(lblSubtitulo);
        Controls.Add(_lblResumen);
        Controls.Add(_dgv);
        Controls.Add(bottomPanel);

        AcceptButton = _btnUnificar;
        CancelButton = _btnCerrar;
    }

    private void LoadDuplicates()
    {
        _binding.Clear();

        var all = _repository.GetAll();

        var withCuit = all
            .Select(c => new
            {
                Company = c,
                CuitDigits = CuitUtils.ExtractDigits(c.CUIT),
                SedeKey = CompanySedeUtils.ComputeSedeKey(c)
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.CuitDigits))
            .ToList();

        var byCuit = withCuit.GroupBy(x => x.CuitDigits).ToList();

        int groupId = 1;
        int duplicateGroups = 0;
        int cuitWithMultipleSedes = 0;

        foreach (var gCuit in byCuit)
        {
            var bySede = gCuit.GroupBy(x => x.SedeKey).ToList();

            if (bySede.Count > 1)
                cuitWithMultipleSedes++;

            foreach (var sedeGroup in bySede)
            {
                var list = sedeGroup.ToList();
                if (list.Count <= 1)
                    continue; // no es duplicado

                duplicateGroups++;

                for (int i = 0; i < list.Count; i++)
                {
                    var item = list[i];
                    var c = item.Company;

                    _binding.Add(new DuplicateRowView
                    {
                        GroupId = groupId,
                        GroupCount = list.Count,
                        CUIT = c.CUIT,
                        Empleador = c.Empleador,
                        Calle = c.Calle,
                        Localidad = c.Localidad,
                        Provincia = c.Provincia,
                        RowIndex = c.RowIndex,
                        Selected = i != 0 // por defecto marcar todos menos el primero para unificar
                    });
                }

                groupId++;
            }
        }

        if (duplicateGroups == 0)
        {
            _lblResumen.Text = "No se encontraron sedes duplicadas (mismo CUIT y misma Calle+Localidad+Provincia).";
        }
        else
        {
            _lblResumen.Text =
                $"Se encontraron {duplicateGroups} sedes duplicadas (mismo CUIT y misma Calle+Localidad+Provincia)." +
                (cuitWithMultipleSedes > 0
                    ? $" También hay {cuitWithMultipleSedes} CUIT con más de una sede (posibles sucursales)."
                    : string.Empty);
        }
    }

    private async Task DeleteSelected(string confirmText)
    {
        var toDelete = _binding.Where(b => b.Selected).ToList();
        if (toDelete.Count == 0)
        {
            MessageBox.Show("No hay registros seleccionados.", "Verificar duplicados",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Asegurar que se mantenga al menos un registro por grupo cuando se usa la opción de unificar.
        var groups = toDelete.GroupBy(d => d.GroupId).ToList();
        foreach (var g in groups)
        {
            var remainingInGroup = _binding.Count(b => b.GroupId == g.Key && !b.Selected);
            if (remainingInGroup == 0)
            {
                MessageBox.Show(
                    $"En el grupo {g.Key} se están intentando eliminar todos los registros. Debe quedar al menos uno.",
                    "Verificar duplicados",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
        }

        var confirm = MessageBox.Show(confirmText,
            "Verificar duplicados",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
            return;

        try
        {
            using (var progress = new ProgressDialog("Verificar duplicados", "Eliminando registros seleccionados..."))
            {
                progress.SetIndeterminate(false);
                progress.Show(this);
                await Task.Yield();

                int total = toDelete.Count;

                await Task.Run(() =>
                {
                    int processed = 0;
                    foreach (var item in toDelete)
                    {
                        var company = _repository.GetByIndex(item.RowIndex);
                        if (company != null)
                        {
                            _repository.DeleteCompany(company);
                        }

                        processed++;
                        var percent = (int)Math.Round(processed * 100.0 / total);
                        progress.SetProgress(percent);
                    }
                });
            }

            LoadDuplicates();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error eliminando registros: {ex.Message}", "Verificar duplicados",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private sealed class DuplicateRowView
    {
        public bool Selected { get; set; }
        public int GroupId { get; set; }
        public int GroupCount { get; set; }
        public string CUIT { get; set; } = string.Empty;
        public string Empleador { get; set; } = string.Empty;
        public string Calle { get; set; } = string.Empty;
        public string Localidad { get; set; } = string.Empty;
        public string Provincia { get; set; } = string.Empty;
        public int RowIndex { get; set; }
    }
}
