let
    OnText = (f as function) as function =>
                (v as any) => if v is text then f(v) else v
in
    [
        Upper       = [Stage = "Pre",  Fn = OnText(Text.Upper)],
        Lower       = [Stage = "Pre",  Fn = OnText(Text.Lower)],
        // схлопує повторні пробіли, нерозривний пробіл вважає звичайним
        CleanSpaces = [Stage = "Pre",  Fn = OnText((t) =>
                         Text.Combine(List.Select(
                             Text.Split(Text.Replace(t, "#(00A0)", " "), " "), each _ <> ""), " "))],
        // прибирає пробіли й нерозривні пробіли в «числах із розрядами» (1 234,56)
        NumText     = [Stage = "Pre",  Fn = OnText((t) => Text.Remove(t, {" ", "#(00A0)", "#(202F)"}))],
        NullToZero  = [Stage = "Post", Fn = (v as any) => v ?? 0]
    ]