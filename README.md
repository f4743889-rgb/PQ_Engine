# Power Query Engine — репозиторій

| Тека | Що в ній | Звідки |
| --- | --- | --- |
| powerquery\ | Код запитів (.pq), queries.yml, таблиці конфігурації (CSV), модулі VBA, README.md з довідкою проєкту | Макрос ExportPowerQueriesToRepo у книзі; вручну не редагувати |
| tests\ | Фікстури регресійних тестів, test_config.xlsx, генератор gen_fixtures.py | Пакет Engine (test_engine, tools) |
| docs\ | Документ архітектури, зовнішній аудит | Пакет Engine (docs) |

Довідка проєкту (правила структури книги, таблиці, запити, макроси): [powerquery/README.md](powerquery/README.md).

Звичайний цикл: зміни в книзі → ExportPowerQueriesToRepo → `git add -A` → `git commit` → `git push`.
Модулі VBA зберігаються в Git в UTF-8 і на диску в cp1251 (.gitattributes); не завантажуйте .bas через вебінтерфейс GitHub.
