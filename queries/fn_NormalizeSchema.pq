// fn_NormalizeSchema
(T0 as table) as table =>
let
    T1   = Table.SelectColumns(T0,
               {"TargetType", "SourceColumnName", "TargetColumnName", "DataType",
                "IsRequired", "TrimFlag", "CustomTransform", "Use"},
               MissingField.UseNull),

    Norm = (v as any) as nullable text =>
               let t = Text.Trim(Text.From(v ?? "")) in if t = "" then null else t,

    CanonT = fn_Canon(cfgSettings[Tables][TargetTypes], "TargetType"),

    T2 = Table.TransformColumns(T1, {
             {"TargetType",       CanonT, type text},
             {"SourceColumnName", Norm, type text},
             {"TargetColumnName", Norm, type text},
             {"DataType",         Norm, type text},
             {"CustomTransform",  Norm, type text}}),

    T3 = Table.SelectRows(T2, each [TargetType] <> null and [SourceColumnName] <> null),

    T4 = Table.TransformColumns(T3, {
             {"IsRequired", each fn_ToBool(_, false), type logical},
             {"TrimFlag",   each fn_ToBool(_, false), type logical},
             {"Use",        each fn_ToBool(_, true),  type logical}})
in
    T4