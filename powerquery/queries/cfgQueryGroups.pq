// cfgQueryGroups: структура тек Power Query (єдине джерело; v4.9)
// Теки в редакторі Power Query розставляються вручну: VBA не має доступу до тек запитів.
// Макрос ExportPowerQueriesToRepo читає цей запит і записує теку кожного запиту в поле group
// файла queries.yml. Новий запит додається сюди (або підпадає під правило Prefixes) і вручну
// переноситься в ту саму теку в редакторі.
// Groups   - тека -> точні імена запитів
// Prefixes - тека для запитів, чиє ім'я починається з префікса (якщо імені немає в Groups)
[
    Groups = [
        #"00 Конфігурація" = {"cfgSettings", "cfgQueryGroups", "cfgTypeMap", "cfgTransforms",
                              "cfgReportTypes", "cfgStatuses", "fn_LoadConfig", "fn_NormalizeSchema",
                              "fn_NormalizeRowFilters", "src_SchemaRegistry"},
        #"01 Утиліти" = {"fn_ToBool", "fn_ToRate", "fn_Opt", "fn_Canon", "fn_TypeKey"},
        #"02 Читання джерел" = {"fn_ActiveSources", "fn_SourcePath", "fn_Options", "fn_ResolveSource",
                                "fn_ParseSource", "fn_ParseHeaders", "fn_ReadSource", "fn_TemplateColumns"},
        #"03 Перевірка налаштувань" = {"fn_RegistryIssues", "pq_RegistryCheck", "pq_SchemaCheck",
                                       "pq_RowFiltersCheck"},
        #"04 Каталог і дрейф схеми" = {"pq_GenerateSchemaTemplate", "pq_ExtractedColumns",
                                       "pq_SchemaIssues", "pq_SchemaDrift"},
        #"05 Нормалізація" = {"fn_GetSchemaRules", "fn_ApplyRowFilters", "fn_ApplyRulesEx", "fn_ApplyRules",
                              "fn_EffectiveSettings", "fn_LoadVerdict", "fn_EnforceSchema"},
        #"06 Якість даних" = {"fn_SchemaReport", "fn_ColumnNullCounts", "fn_RowReport", "fn_QcReports", "fn_SummaryTable",
                              "pq_QualitySummary", "pq_QualityColumns", "pq_QualityErrors"},
        #"07 Призначення" = {"fn_ReadPurpose"},
        #"08 Модель і звіт" = {"Calendar", "fxReadNamedCell", "targetReportCell"}
    ],
    Prefixes = [
        q_    = "07 Призначення",
        test_ = "99 Тести"
    ]
]