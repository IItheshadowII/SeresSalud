using ConvertidorDeOrdenes.Desktop.Services;

namespace ConvertidorDeOrdenes.Desktop.Forms;

public sealed class ProgressDialog : Form
{
    private readonly Label _lbl = null!;
    private readonly ProgressBar _bar = null!;

    public ProgressDialog(string title, string initialMessage)
    {
        Text = title;
        Icon = AppIcon.TryGet();
        Size = new Size(520, 160);
        // CenterParent no centra correctamente cuando se usa Show() (modeless).
        // Lo centramos manualmente en OnShown.
        StartPosition = FormStartPosition.Manual;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        BackColor = Color.WhiteSmoke;

        _lbl = new Label
        {
            Location = new Point(20, 20),
            Size = new Size(460, 40),
            Font = new Font("Segoe UI", 10, FontStyle.Regular),
            ForeColor = Color.FromArgb(60, 60, 60),
            Text = initialMessage
        };

        _bar = new ProgressBar
        {
            Location = new Point(20, 75),
            Size = new Size(460, 22),
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30
        };

        Controls.Add(_lbl);
        Controls.Add(_bar);
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);

        try
        {
            Rectangle bounds;

            if (Owner != null)
            {
                bounds = Owner.Bounds;
            }
            else
            {
                var screen = Screen.FromPoint(Cursor.Position);
                bounds = screen.WorkingArea;
            }

            var x = bounds.Left + (bounds.Width - Width) / 2;
            var y = bounds.Top + (bounds.Height - Height) / 2;

            // Evitar quedar fuera de pantalla
            var wa = Screen.FromRectangle(bounds).WorkingArea;
            x = Math.Max(wa.Left, Math.Min(x, wa.Right - Width));
            y = Math.Max(wa.Top, Math.Min(y, wa.Bottom - Height));

            Location = new Point(x, y);
        }
        catch
        {
            // No bloquear si falla el centrado.
        }
    }

    public void SetStatus(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetStatus(message));
            return;
        }

        _lbl.Text = message;
    }

    public void SetIndeterminate(bool indeterminate)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetIndeterminate(indeterminate));
            return;
        }

        _bar.Style = indeterminate ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
        if (indeterminate)
        {
            _bar.MarqueeAnimationSpeed = 30;
        }
        else
        {
            _bar.MarqueeAnimationSpeed = 0;
            _bar.Minimum = 0;
            _bar.Maximum = 100;
            _bar.Value = 0;
        }
    }

    public void SetProgress(int percent)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetProgress(percent));
            return;
        }

        if (_bar.Style == ProgressBarStyle.Marquee)
            return;

        percent = Math.Max(0, Math.Min(100, percent));
        _bar.Value = percent;
    }
}
