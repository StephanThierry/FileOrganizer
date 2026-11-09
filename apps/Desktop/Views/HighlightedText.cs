using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;

namespace FileOrganizer.Desktop.Views;

/// <summary>Text where the parts that matched are highlighted (the regex guide's test box).</summary>
public class HighlightedText : SelectableTextBlock
{
    public static readonly StyledProperty<IReadOnlyList<TextPiece>?> PiecesProperty =
        AvaloniaProperty.Register<HighlightedText, IReadOnlyList<TextPiece>?>(nameof(Pieces));

    public IReadOnlyList<TextPiece>? Pieces
    {
        get => GetValue(PiecesProperty);
        set => SetValue(PiecesProperty, value);
    }

    protected override Type StyleKeyOverride => typeof(SelectableTextBlock);

    static HighlightedText() => PiecesProperty.Changed.AddClassHandler<HighlightedText>((t, _) => t.Rebuild());

    private void Rebuild()
    {
        Inlines?.Clear();
        if (Pieces == null) return;
        Inlines ??= new InlineCollection();
        // The theme's accent colour, see-through, so the text stays readable in the light and the dark theme
        var accent = this.TryFindResource("Accent", ActualThemeVariant, out var found) && found is ISolidColorBrush solid ? solid.Color : Colors.Gold;
        foreach (var p in Pieces)
        {
            var run = new Run(p.Text);
            if (p.IsMatch)
            {
                run.Background = new SolidColorBrush(accent, 0.35);
                run.FontWeight = FontWeight.Bold;
            }
            Inlines.Add(run);
        }
    }
}
