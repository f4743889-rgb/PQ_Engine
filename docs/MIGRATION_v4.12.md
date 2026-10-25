# Перехід з v4.11 на v4.12

Що змінилося і чому — `docs/architecture.md`, розділ 6.12 («Заміри 25.10.2026») і журнал змін. Перехід з v4.10 — `MIGRATION_v4.11.md` (спершу він, потім цей).

## 1. Робоча книга

1. **Запити.** ImportPQQueriesFromFolder → `for_workbook\powerquery`. Змінилися pq_QualitySummary, pq_QualityColumns, pq_QualityErrors (тепер самі викликають fn_QcReports(fn_LoadConfig())), а також коментарі в cfgQueryGroups, fn_LoadConfig, fn_NormalizeRowFilters, fn_QcReports.
2. **Видаліть запити** qc_Reports і тимчасові pq_QualitySummary2, pq_QualitySummary3, fn_SchemaReport_nb, fn_RowReport_nb, fn_QcReports_nb (і їхні аркуші). Імпорт сам нічого не видаляє.
3. **VBA.** Замініть modPOWERQUERYPROFILER (колонка Seq, живлення в Notes) і modTableCatalog (порядок правил hlpProject, стартовий опис pq_ExtractedColumns). Debug → Compile.
4. Оновити все → pq_QualitySummary має ті самі рядки, що й раніше. Formula.Firewall не очікується: так само побудований test_Summary, а pq_QualitySummary2 уже працював у робочій книзі.
5. UpdateHelp (hlpProject оновиться: розділ «Довідка» — три правила поспіль), потім ExportPowerQueriesToRepo і коміт. У Git: видалений qc_Reports.pq, три змінені звіти, profiler, catalog, queries.yml, README.

## 2. Тестова книга

Ті самі модулі; SetupTestEngine з кореня клону (після коміту робочої книги) або з теки пакета. Видаліть qc_Reports, якщо лишився. Оновити все → test_All: 11 PASS.

## 3. Рішення за вами

- **Calendar:** зняти «Оновлювати під час оновлення всіх» і оновлювати вручну після зміни tblPeriod (−4…7,6 с на кожному «Оновити все»).
- **R3:** виключити звіти якості з «Оновити все» й оновлювати через RefreshWithRunLog (−25…35 с на кожному звичайному оновленні).

## 4. Замір (необов'язково)

Профайлер від мережі, режим «Найкраща продуктивність», 2–3 запуски. Очікування: pq_QualitySummary ≈ колишній pq_QualitySummary2 (~25–35 с).
