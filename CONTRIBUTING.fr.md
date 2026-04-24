# Contribuer à ExcelHelper

Merci de votre intérêt pour contribuer à ExcelHelper !

## Pour commencer

1. Forker le dépôt
2. Cloner votre fork
3. Créer une branche de fonctionnalité (`git checkout -b feature/ma-fonctionnalite`)
4. Construire la solution : `dotnet build ExcelHelper.sln`
5. Exécuter les tests : `dotnet test ExcelHelper.sln`

## Prérequis

- Tout le code doit compiler sur **tous les frameworks cibles** (`net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0`)
- **Zéro avertissement** : `TreatWarningsAsErrors` est activé
- **Documentation XML** requise pour toutes les API publiques
- **Tests unitaires** requis pour toutes les nouvelles fonctionnalités
- Suivre le style de code existant

## Normes de codage

1. Utiliser les annotations `Nullable` (`?`) et les clauses de garde
2. Utiliser `CompiledExpressionCache` pour les accès get/set de propriétés dans les chemins critiques
3. Minimiser les allocations dans les boucles
4. Supporter l'API Fluent (chaînage de méthodes)
5. Utiliser `IAsyncEnumerable<T>` pour .NET Core+, `Task<IReadOnlyList<T>>` pour .NET Framework

## Processus de demande de tirage (PR)

1. Chaque PR doit cibler une seule phase ou fonctionnalité
2. La construction doit passer sur tous les frameworks cibles (`dotnet build`)
3. Tous les tests doivent passer (`dotnet test`)
4. La documentation XML doit être complète
5. Mettre à jour `ARCHITECTURE.md` si vous ajoutez de nouveaux points d'extension
6. Au moins une approbation d'un relecteur est requise

## Signalement de problèmes

Veuillez inclure :
- Framework cible
- Version d'ExcelHelper
- Code de reproduction minimal
- Comportement attendu vs comportement actuel
