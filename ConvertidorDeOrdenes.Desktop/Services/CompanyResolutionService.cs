using ConvertidorDeOrdenes.Core.Models;
using ConvertidorDeOrdenes.Core.Services;
using ConvertidorDeOrdenes.Desktop.Forms;

namespace ConvertidorDeOrdenes.Desktop.Services;

public sealed class CompanyResolutionService
{
    private readonly CompanyRepositoryExcel _repository;
    private readonly CompanyResolutionDecisionStore _store;

    public CompanyResolutionService(CompanyRepositoryExcel repository)
    {
        _repository = repository;
        _store = new CompanyResolutionDecisionStore(AppPaths.CompanyResolutionDecisionsPath);
    }

    /// <summary>
    /// Guarda una empresa en la base aplicando la política CUIT/sede:
    /// - Si existe mismo CUIT: mostrar como "Posible sede adicional".
    /// - Solo considerar "duplicado" si coincide sede.
    /// - Acciones: unificar / mantener / ignorar.
    /// - Persistir decisión para no repetir.
    /// </summary>
    public bool SaveWithResolution(IWin32Window owner, CompanyRecord company)
    {
        if (company == null)
            throw new ArgumentNullException(nameof(company));

        company.CUIT = CuitUtils.FormatOrKeep(company.CUIT);

        // Edición explícita: si tiene RowIndex, guardar directo.
        if (company.RowIndex > 0)
        {
            _repository.SaveCompany(company, forceNew: false);
            return true;
        }

        var cuitDigits = CuitUtils.ExtractDigits(company.CUIT);
        if (string.IsNullOrWhiteSpace(cuitDigits))
        {
            _repository.SaveCompany(company, forceNew: false);
            return true;
        }

        var existingByCuit = _repository.SearchByCuit(company.CUIT);
        if (existingByCuit.Count == 0)
        {
            _repository.SaveCompany(company, forceNew: false);
            return true;
        }

        var sedeKey = CompanySedeUtils.ComputeSedeKey(company);

        if (_store.TryGet(company.CUIT, sedeKey, out var cached))
        {
            var applied = ApplyDecision(company, existingByCuit, cached);
            if (applied)
                return true;
            // Si no se pudo aplicar (ej: rowIndex ya no existe), caer a interacción.
        }

        using var dialog = new CompanyCuitConflictDialog(company, existingByCuit);
        if (dialog.ShowDialog(owner) != DialogResult.OK)
            return false;

        var decision = new CompanyResolutionDecision
        {
            Kind = dialog.Decision,
            TargetRowIndex = dialog.SelectedTargetRowIndex
        };

        var saved = ApplyDecision(company, existingByCuit, decision);
        _store.Save(company.CUIT, sedeKey, decision);
        return saved;
    }

    private bool ApplyDecision(
        CompanyRecord company,
        List<CompanyRecord> existingByCuit,
        CompanyResolutionDecision decision)
    {
        switch (decision.Kind)
        {
            case CompanyResolutionDecisionKind.Unify:
            {
                var targetRow = decision.TargetRowIndex;
                if (targetRow == null)
                    return false;

                var exists = existingByCuit.Any(c => c.RowIndex == targetRow.Value);
                if (!exists)
                {
                    // Si el registro ya no existe, no asumir; dejar que el flujo vuelva a preguntar.
                    return false;
                }

                company.RowIndex = targetRow.Value;
                _repository.SaveCompany(company, forceNew: false);
                return true;
            }

            case CompanyResolutionDecisionKind.Keep:
                company.RowIndex = 0;
                _repository.SaveCompany(company, forceNew: true);
                return true;

            case CompanyResolutionDecisionKind.Ignore:
                return false;

            default:
                return false;
        }
    }
}
