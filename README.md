# Power Query Engine — репозиторій

| Тека | Що в ній | Звідки |
| --- | --- | --- |
| powerquery\ | Код запитів (.pq), queries.yml, таблиці конфігурації (CSV), модулі VBA, README.md з довідкою проєкту | Макрос ExportPowerQueriesToRepo **у робочій книзі**; вручну не редагувати |
| tests\ | Фікстури, test_config.xlsx, тестові запити test_*.pq, генератор gen_fixtures.py | Пакет Engine; оновлюються вручну разом зі змінами тестів |
| docs\ | Документ архітектури (architecture.md), зовнішній аудит | Пакет Engine |

Довідка проєкту (правила структури книги, таблиці, запити, макроси): [powerquery/README.md](powerquery/README.md).

**Яка книга дзеркалиться.** Тека powerquery\ — це робоча книга (лежить у корені клону, у Git не потрапляє).
Тестову книгу тримають поза репозиторієм; ExportPowerQueriesToRepo у ній зупиняється, щоб тестовий реєстр і схема
не замінили робочі. Код запитів Engine в обох книгах однаковий; тестова конфігурація — у tests\test_config.xlsx.

Звичайний цикл: зміни в робочій книзі → ExportPowerQueriesToRepo → GitHub Desktop: перегляд змін → commit → push.
Модулі VBA зберігаються в Git в UTF-8 і на диску в cp1251 (.gitattributes); не завантажуйте .bas через вебінтерфейс GitHub.
