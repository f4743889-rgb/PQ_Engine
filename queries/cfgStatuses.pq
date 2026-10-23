// cfgStatuses: довідник статусів зведення якості (аудит D5)
// Order    - порядок у pq_QualitySummary (проблеми вгорі)
// Blocking - статус сам по собі означає, що fn_ReadPurpose зупиниться (QualityFlag = FAIL);
//            для ERROR рішення приймає fn_LoadVerdict (обов'язкові колонки, політика, поріг)
[
    READ_ERROR = [Order = 1, Blocking = true,  Text = "Файл знайдено, але прочитати чи обробити його не вдалося (повідомлення в Message)"],
    FILE_ERROR = [Order = 2, Blocking = true,  Text = "Обов'язковий файл недоступний чи шлях порожній"],
    NO_SCHEMA  = [Order = 3, Blocking = true,  Text = "Для TargetType немає правил або вони некоректні"],
    ERROR      = [Order = 4, Blocking = false, Text = "Немає обов'язкової колонки або є помилки типізації"],
    WARN       = [Order = 5, Blocking = false, Text = "Немає необов'язкових колонок, є повністю порожні або дубль колонки схеми"],
    EMPTY      = [Order = 6, Blocking = false, Text = "Файл прочитано, але даних немає"],
    NODATA     = [Order = 7, Blocking = false, Text = "Необов'язкового файла немає (чи шлях не заданий)"],
    OK         = [Order = 8, Blocking = false, Text = "Усе гаразд"]
]