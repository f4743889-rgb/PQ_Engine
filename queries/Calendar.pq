let
    Source = Excel.CurrentWorkbook(){[Name = cfgSettings[Tables][Period]]}[Content],

    Period = Table.TransformColumnTypes(
        Source,
        {
            {"StartDate", type date},
            {"EndDate", type date}
        }
    ),

    StartDate = List.Min(Period[StartDate]),
    EndDate = List.Max(Period[EndDate]),

    Dates = List.Dates(
        StartDate,
        Duration.Days(EndDate - StartDate) + 1,
        #duration(1, 0, 0, 0)
    ),

    Calendar = Table.FromColumns(
        {Dates},
        type table [Date = date]
    ),

    AddYear = Table.AddColumn(
        Calendar, "Рік",
        each Date.Year([Date]), Int64.Type
    ),

    AddQuarter = Table.AddColumn(
        AddYear, "Номер кварталу",
        each Date.QuarterOfYear([Date]), Int64.Type
    ),

    AddMonthNumber = Table.AddColumn(
        AddQuarter, "Номер місяця",
        each Date.Month([Date]), Int64.Type
    ),

    AddMonthName = Table.AddColumn(
        AddMonthNumber, "Місяць",
        each Date.MonthName([Date], cfgSettings[Culture]), type text
    ),

    AddYearMonth = Table.AddColumn(
        AddMonthName, "Рік-місяць",
        each Date.ToText([Date], "yyyy-MM", cfgSettings[Culture]), type text
    )
in
    AddYearMonth