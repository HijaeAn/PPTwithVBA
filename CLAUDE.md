# CLAUDE.md

his file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
dotnet build
dotnet run
```

Target framework: `net8.0-windows`. No test project exists yet.

## Code Style Guide

### Naming
- Variable names: always **camelCase** (e.g., `filteredCoinList`, `priceChangePercent`)
- Avoid single-letter names, abbreviations, or generic names like `data`, `temp`, `val`
- Use names that clearly communicate purpose

### Syntax
- 람다식보다 일반적인 코드 구문으로 작성할 것

## Architecture

MVVM pattern — no third-party MVVM framework, only built-in WPF bindings.

- **`ViewModels/MainViewModel.cs`** — single `bool IsOn` property (`INotifyPropertyChanged`) drives all UI state. `ToggleCommand` flips it.
- **`ViewModels/RelayCommand.cs`** — minimal `ICommand` wrapper delegating `CanExecuteChanged` to `CommandManager.RequerySuggested`.
- **`Converters/BoolToVisibilityConverters.cs`** — two converters in one file: `BoolToVisibilityConverter` (true→Visible) and `InverseBoolToVisibilityConverter` (true→Collapsed). Used for group A/B visibility switching.
- **`MainWindow.xaml`** — `DataContext` set directly in XAML via `<vm:MainViewModel />`. Resources declare both toggle styles and converter instances. No code-behind logic.

## Key Patterns

**Toggle button styles** — two named styles in `Window.Resources`, both bound to the same `IsOn`:
- `OnOffToggleStyle` — text label button, color via `Style.Triggers`
- `SlidingToggleStyle` — iOS-style switch using `ControlTemplate` with `ColorAnimation` + `DoubleAnimation` (0.15 s) on `IsChecked` triggers

**Visibility control** — group A uses `BoolToVis`, group B uses `InverseBoolToVis`. Both grids are declared `Visible` in XAML (design-time friendly); converters take over at runtime. Uses `Collapsed` (not `Hidden`) so hidden groups don't occupy space.

**Binding direction** — both toggle buttons use `Mode=TwoWay` on `IsChecked`, so clicking either one updates the shared `IsOn` and the other button reflects the change immediately.
