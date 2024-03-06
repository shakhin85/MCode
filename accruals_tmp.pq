let
     f = (x as text, report_date as text)=>
     [
         a = Excel.Workbook(File.Contents(
             x
         )),
         b = a{[Name = "для IF загрузки"]}[Data],
         
                    func = (table, optional options)=>

                    [ tbl = Table.Buffer(Table.Profile(table)),
                      lst = Table.SelectRows(tbl, (r)=>r[Count]=r[NullCount])[Column],
                      to = Table.RemoveColumns(table, lst)][to]
                    ,
          c = func(b),
          Added_index_col = Table.AddIndexColumn(c,"Index"),
          #"Changed Type" = Table.TransformColumnTypes(Added_index_col,{{"Index", type text}}),
          Header_rownumber = Int16.From(Table.FindText(#"Changed Type","Отдел инициатор")[Index]{0}),
          cols_to_be = List.RemoveLastN(Record.FieldValues(#"Changed Type"{Header_rownumber}),1),
          cols_as_is = Table.ColumnNames(c),
          zipped = List.Zip({cols_as_is, cols_to_be}),
          renamed = Table.RenameColumns(c, zipped),
          to_1 = Table.Range(renamed, Header_rownumber+1),
          #"define non-null rows" = Table.AddColumn(to_1, "Param", each 
          List.NonNullCount(List.RemoveMatchingItems(Record.FieldValues(_), {null}))),
          #"Filtered Rows" = Table.SelectRows(#"define non-null rows", each ([Param] > 1)),
          #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Param"}),
        dates_to_text = List.Accumulate(Table.ColumnNames(#"Removed Columns"),
            #"Removed Columns", (s,c)=> if c = "Дт" or c = "Кт" then Table.TransformColumns(
                s, {c, each if Value.Type(_) = Date.Type then Date.ToText(_, "dd.MM.yy")
                    else if _="08.01.2004" then "08.01.04"
                    else if _="04.01.2001" then "04.01.01"
                        
                        else _}) else s
            ),
        to_11 = Table.TransformRows(dates_to_text, (x) =>
            Record.TransformFields(x, {"МВЗ",
            each if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"Mobility","Мобилити","Lab_файл"}, (value) => Text.Contains(x[#"Отдел инициатор"], value))) then 99013
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"Genesis"}, (value) => Text.Contains(x[#"Отдел инициатор"], value))) then 99001
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"HR"}, (value) => Text.Contains(x[#"Отдел инициатор"], value))) then 11001
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"PR"}, (value) => Text.Contains(x[#"Отдел инициатор"], value))) then 99004
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"HTA"}, (value) => Text.Contains(x[#"Отдел инициатор"], value))) then 99002
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"Административный"}, (value) => Text.StartsWith(x[#"Отдел инициатор"], value))) then 11002
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"отдел послепродажного обслуживания"}, (value) => Text.StartsWith(x[#"Отдел инициатор"], value))) then 55001
            else if Value.Is(_, Null.Type) and List.AnyTrue(List.Transform({"Бухгалтерия","Финансовый отдел","отдел реализации"}, (value) => Text.StartsWith(x[#"Отдел инициатор"], value))) then 66001

            
            else _})
        ),
        to_2 = Table.FromRecords(to_11),
        #"Added Custom" = Table.AddColumn(to_2, "Source", each x),
        
        #"Trimmed Text" = Table.TransformColumns(#"Added Custom",{{"Дт", Text.Trim, type text},{"Кт", Text.Trim, type text}}),
        Buffered_Mapping = Table.Buffer(
          Expression.Evaluate(     
          Text.FromBinary(Web.Contents(
          "https://raw.githubusercontent.com/shakhin85/MCode/main/Mapping.pq"
               )), #shared)

               ), 
        Custom0 = Table.ToRecords(Buffered_Mapping),    
        Custom1 = Table.AddColumn(#"Trimmed Text", "Dr Corp Account", each List.Select( Custom0,(x)=>x[Rus Account]=[#"Дт"]){0}?[#"Local Account Code"]?),
        Custom2 = Table.AddColumn(Custom1, "Cr Corp Account", each List.Select( Custom0,(x)=>x[Rus Account]=[#"Кт"]){0}?[#"Local Account Code"]?),
        Custom3 = Table.AddColumn(Custom2, "Dr Acc_group", each List.Select( Custom0,(x)=>x[Rus Account]=[#"Дт"]){0}?[#"Balance or PnL"]?),
        Custom4 = Table.AddColumn(Custom3, "Сr Acc_group", each List.Select( Custom0,(x)=>x[Rus Account]=[#"Кт"]){0}?[#"Balance or PnL"]?),
        #"Added Custom1" = Table.AddColumn(Custom4, "PL param", each if [Dr Acc_group] = "PL" and [Сr Acc_group] <> "PL" then "pl"
                                                                else if [Dr Acc_group] <> "PL" and [Сr Acc_group] = "PL"
                                                                then "pl"
                                                                else if [Dr Acc_group] = "PL" and [Сr Acc_group] = "PL"
                                                                then "pl"
                                                                else "bs"),
        #"Filtered Rows1" = Table.SelectRows(#"Added Custom1", each ([PL param] = "pl")),
        #"Changed Type1" = Table.TransformColumnTypes(#"Filtered Rows1",{{"Сумма", type number}}),
        #"Removed Columns1" = Table.RemoveColumns(#"Changed Type1",{"Период, к которому относятся расходы"}),
        #"added_report_date" = Table.AddColumn(#"Removed Columns1","Reporting date", each
        Date.FromText(report_date), type date),

        #"Added Custom2" = Table.AddColumn(#"added_report_date", "Corp Acc",
        each if List.AnyTrue(
                List.Transform({"1", "2", "3"}, 
                    (prefix) => Text.StartsWith(Text.From([Dr Corp Account]), prefix)
                )
            ) or [Dr Corp Account] is null 
        then [Cr Corp Account] 
        else [Dr Corp Account]),
        added_sign_control = Table.AddColumn(#"Added Custom2", "Sign control",
        each List.Select( Custom0,(x)=>x[Local Account Code]=[#"Corp Acc"]){0}?[#"Sign control"]?),
        final = Table.FromRecords(
        Table.TransformRows(added_sign_control, (r)=>
        Record.TransformFields(r, {"Сумма", 
        each if r[Сумма]>0  and r[#"Sign control"]=-1
        then _*-1
        else if r[Сумма]>0 and List.AnyTrue(List.Transform(#"FX accounts"[#"Local Account Code"],
                                (value) => r[#"Dr Corp Account"] = value))
        then _*1
        else if r[Сумма]>0 and List.AnyTrue(List.Transform(#"FX accounts"[#"Local Account Code"],
                                (value) => r[#"Cr Corp Account"] = value))
        then _*-1
        
        else _
        
        }))
        ),
        #"Changed Type2" = Table.TransformColumnTypes(final,{{"Сумма", type number}}),
        #"Removed Columns2" = Table.RemoveColumns(#"Changed Type2",{"Dr Corp Account", "Cr Corp Account", "Dr Acc_group", "Сr Acc_group", "PL param", "Sign control", "НОМЕР ROA"}),
        Company_code_added = Table.AddColumn(#"Removed Columns2","Company_code",
        each if [#"Кредитор/дебитор ДТ"] is null then [#"Кредитор/дебитор КТ"] else [#"Кредитор/дебитор ДТ"]),
        #"Added Custom3" = List.Accumulate(
            {"MonthNumber","Year"}, Company_code_added, (s,c)=>if c = "MonthNumber" then
            Table.AddColumn(s, c, each Date.Month([Reporting date])) else Table.AddColumn(s, c, each Date.Year([Reporting date]))
        ),
        #"Removed Columns3" = Table.RemoveColumns(#"Added Custom3",{"Кредитор/дебитор ДТ", "Кредитор/дебитор КТ", "Код НДС", "Дт", "Кт"}),
        selected_cols = Table.SelectColumns(#"Removed Columns3",
        {"Year","MonthNumber","Corp Acc","Отдел инициатор","Ответственный отдел в ОБО",
        "Описание операции","Сумма","МВЗ","Reporting date","Company_code"}, MissingField.Ignore),
        Renamed_cols_to_comply_with_PL_data = Table.RenameColumns(
            selected_cols, List.Zip({Table.ColumnNames(selected_cols), 
            {
        "Year","MonthNumber","GL_Acc","Cost center name",
        "Doc number","Description text","Net of","Cost Center code","Reporting date","Company code"
    }}), MissingField.Ignore
        ),
        #"Added Custom4" = Table.AddColumn(Renamed_cols_to_comply_with_PL_data, "Source", each "Accruals"),
        #"Added Custom5" = Table.AddColumn(#"Added Custom4", "MMYYYY", each Text.Upper(Date.ToText([Reporting date], "MMM ","EN-en")) & Number.ToText([Year])),
        final_adj =Table.TransformColumns(#"Added Custom5", 
        {"MonthNumber", each if Text.StartsWith(Number.ToText(_), "0") then Text.Replace(_,"0","") else Number.ToText(_), type text}),
        #"Changed Type3" = Table.TransformColumnTypes(final_adj,
        {{"Reporting date", type date}, {"Cost Center code", type text}, {"Company code", type text}, {"Source", type text}, {"MMYYYY", type text}, {"Year", type text}, {"GL_Acc", type text},
        {"Cost center name", type text}, {"Doc number", type text}, {"Description text", type text}}),
        final_table = Table.FromRecords(
        Table.TransformRows(#"Changed Type3", (row)=>Record.TransformFields(row,
        {
            "Net of", each if row[GL_Acc] = "41010001" and row[Source] = "Accruals" then _ * -1 else _
        }))
        )][final_table]
     
in
    f
