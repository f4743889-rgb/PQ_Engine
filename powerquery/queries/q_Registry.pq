// q_Registry: запит призначення Registry (зразок для інших q_...)
// Конфігурація (реєстр, схема, фільтри службових рядків) - fn_LoadConfig.
let
    Result = fn_ReadPurpose(fn_LoadConfig(), "Registry")
in
    Result