// fn_ToBool: Гнучке приведення до logical

(v as any, default as logical) as logical =>
    if v = null then default
    else if v is logical then v
    else if v is number then v <> 0
    else
        let t = Text.Lower(Text.Trim(Text.From(v)))
        in if t = "" then default else List.Contains(cfgSettings[TrueValues], t)