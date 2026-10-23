// fn_ActiveSources: єдине визначення «активного» джерела (Enabled; порожнє = TRUE);
// TargetType і PeriodType приводяться до написання довідників tblTargetType, tblPeriodRules
(registry as table) as table =>
let
    IsBlankRow = (r as record) as logical =>
        List.AllTrue(List.Transform(Record.FieldValues(r),
            (v) => v = null or (v is text and Text.Trim(v) = ""))),
    NotBlank = Table.SelectRows(registry, each not IsBlankRow(_)),
    Active   = Table.SelectRows(NotBlank,
                   each fn_ToBool(Record.FieldOrDefault(_, "Enabled", null), true)),

    // правило регістру: TargetType і PeriodType у написанні довідників (fn_Canon)
    Canon    = Table.TransformColumns(Active,
                   {{"TargetType", fn_Canon(cfgSettings[Tables][TargetTypes], "TargetType")},
                    {"PeriodType", fn_Canon(cfgSettings[Tables][PeriodRules], "PeriodType")}},
                   null, MissingField.Ignore)
in
    Canon