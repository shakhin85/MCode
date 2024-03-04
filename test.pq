let
    // Создаем фиктивные таблицы для трех месяцев с дополнительными атрибутами
    Acc_1m = Table.FromRecords({
        [DocumentName="Doc1", Department="Sales", OperationText="Sale", Initiator="John", Month=1, Year = 2024, #"Cumulative Amount"=100],
        [DocumentName="Doc2", Department="HR", OperationText="Hire", Initiator="Alice", Month=1, Year = 2024, #"Cumulative Amount"=200]
    }),
    Acc_2m = Table.FromRecords({
        [DocumentName="Doc1", Department="Sales", OperationText="Sale", Initiator="John", Month=2, Year = 2024, #"Cumulative Amount"=150],
        [DocumentName="Doc2", Department="HR", OperationText="Hire", Initiator="Alice", Month=2, Year = 2024, #"Cumulative Amount"=250]
    }),
    Acc_3m = Table.FromRecords({
        [DocumentName="Doc1", Department="Sales", OperationText="Sale", Initiator="John", Month=3, Year = 2024, #"Cumulative Amount"=200],
        [DocumentName="Doc2", Department="HR", OperationText="Hire", Initiator="Alice", Month=3, Year = 2024, #"Cumulative Amount"=300]
    }),

    // Создаем список таблиц для удобства итерации
    TablesList = {Acc_1m, Acc_2m, Acc_3m},
    CombinedTables = Table.Combine(TablesList),

    // Добавляем композитный ключ
    AddedCompositeKey = Table.AddColumn(CombinedTables, "Composite Key", each Text.Combine({[DocumentName], [Department], [OperationText], [Initiator]}, "|")),

    // Сортируем таблицу по композитному ключу и месяцу
    SortedTable = Table.Sort(AddedCompositeKey, { {"Month", Order.Ascending},{"Composite Key", Order.Ascending}  }),

    // Группируем отсортированные данные по композитному ключу
    GroupedTable = Table.Group(SortedTable, {"Composite Key"}, {
        {"AllData", each _, type table [DocumentName=text, Department=text, OperationText=text, Initiator=text, MonthYear=text, #"Cumulative Amount"=number,
        #"Composite Key"=text]}
    }),

    // Вычисляем месячные суммы для каждой группы
    TransformedTable = Table.TransformColumns(GroupedTable, {
        "AllData", (t) => let
            IndexedTable = Table.AddIndexColumn(t, "Index", 1, 1, Int64.Type),
            MonthlyAmounts = Table.FromRecords(
                List.Transform(
                    Table.ToRecords(IndexedTable),
                    (currentRecord) =>
                        let
                            PrevRecord = if currentRecord[Index] > 1 then IndexedTable{currentRecord[Index] - 2} else null,
                            PrevAmount = if PrevRecord = null then 0 else PrevRecord[#"Cumulative Amount"],
                            MonthlyAmount = currentRecord[#"Cumulative Amount"] - PrevAmount
                        in
                            Record.AddField(currentRecord, "Monthly Amount",  MonthlyAmount)
                )
            )
        in
            MonthlyAmounts
    }),
    #"Expanded AllData" = Table.ExpandTableColumn(TransformedTable, "AllData", {"DocumentName", "Department", "OperationText", "Initiator", "Month", "Year", "Cumulative Amount", "Index", "Monthly Amount"}, {"DocumentName", "Department", "OperationText", "Initiator", "Month", "Year", "Cumulative Amount", "Index", "Monthly Amount"})
in
    #"Expanded AllData"
