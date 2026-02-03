using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

public sealed class CompanyCuitConflictDialog : Form
{
    private readonly CompanyRecord _incoming;
    private readonly List<CompanyRecord> _existing;

    private ListBox _lstExisting = null!;
    private Label _lblHeader = null!;
    private Label _lblHint = null!;

    public CompanyResolutionDecisionKind Decision { get; private set; }
    public int? SelectedTargetRowIndex { get; private set; }

    public CompanyCuitConflictDialog(CompanyRecord incoming, List<CompanyRecord> existing)
    {
        _incoming = incoming;
        _existing = existing;

        InitializeComponent();
        LoadExisting();
    }

    private void InitializeComponent()
    {
        Text = "Mismo CUIT detectado";
        Icon = AppIcon.TryGet();
        Size = new Size(760, 520);
        MinimumSize = new Size(760, 520);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.WhiteSmoke;

        _lblHeader = new Label
        {
            Location = new Point(20, 15),
            Size = new Size(710, 46),
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };

        _lblHint = new Label
        {
            Location = new Point(20, 62),
            Size = new Size(710, 36),
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90)
        };

        _lstExisting = new ListBox
        {
            Location = new Point(20, 105),
            Size = new Size(710, 270),
            Font = new Font("Segoe UI", 9),
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };

        var btnUnify = new Button
        {
            Text = "✅ Es la misma sede (unificar)",
            Location = new Point(20, 395),
            Size = new Size(230, 42),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnUnify.FlatAppearance.BorderSize = 0;
        btnUnify.Click += (_, _) => Choose(CompanyResolutionDecisionKind.Unify);

        var btnKeep = new Button
        {
            Text = "➕ Es otra sede (mantener)",
            Location = new Point(260, 395),
            Size = new Size(230, 42),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(60, 140, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnKeep.FlatAppearance.BorderSize = 0;
        btnKeep.Click += (_, _) => Choose(CompanyResolutionDecisionKind.Keep);

        var btnIgnore = new Button
        {
            Text = "🚫 No estoy seguro (ignorar por ahora)",
            Location = new Point(500, 395),
            Size = new Size(230, 42),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(160, 120, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnIgnore.FlatAppearance.BorderSize = 0;
        btnIgnore.Click += (_, _) => Choose(CompanyResolutionDecisionKind.Ignore);

        Controls.Add(_lblHeader);
        Controls.Add(_lblHint);
        Controls.Add(_lstExisting);
        Controls.Add(btnUnify);
        Controls.Add(btnKeep);
        Controls.Add(btnIgnore);

        AcceptButton = btnUnify;
    }

    private void LoadExisting()
    {
        var formatted = CuitUtils.FormatOrKeep(_incoming.CUIT);
        var incomingSede = CompanySedeUtils.ComputeSedeKey(_incoming);

        bool anySameSede = _existing.Any(e =>
            CompanySedeUtils.ComputeSedeKey(e).Equals(incomingSede, StringComparison.OrdinalIgnoreCase));

        _lblHeader.Text = anySameSede
            ? $"Duplicado: mismo CUIT y misma sede ({formatted})."
            : $"Posible sede adicional: ya existe el CUIT {formatted}.";

        _lblHint.Text = "Solo se marca como duplicado si coincide la sede. Seleccione un registro para unificar o elija mantener.";

        _lstExisting.Items.Clear();

        int selectIndex = 0;
        for (int i = 0; i < _existing.Count; i++)
        {
            var c = _existing[i];
            var sedeKey = CompanySedeUtils.ComputeSedeKey(c);
            var sedeLabel = sedeKey.Equals(incomingSede, StringComparison.OrdinalIgnoreCase) ? " (misma sede)" : "";

            _lstExisting.Items.Add($"[{c.RowIndex}] {c.Empleador} — {c.Calle} — {c.CodPostal} {c.Localidad}, {c.Provincia}{sedeLabel}");

            if (sedeLabel.Length > 0)
                selectIndex = i;
        }

        if (_lstExisting.Items.Count > 0)
            _lstExisting.SelectedIndex = selectIndex;
    }

    private void Choose(CompanyResolutionDecisionKind kind)
    {
        Decision = kind;

        if (kind == CompanyResolutionDecisionKind.Unify)
        {
            if (_lstExisting.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione una empresa existente para unificar.", "Validación",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SelectedTargetRowIndex = _existing[_lstExisting.SelectedIndex].RowIndex;
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
