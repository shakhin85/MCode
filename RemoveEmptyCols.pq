let
    Source = [


func = (table, optional options)=>

[
    tbl = Table.Buffer(Table.Profile(table)),
    lst = Table.SelectRows(tbl, (r)=>r[Count]=r[NullCount])[Column],
    to = Table.RemoveColumns(table, lst)] [to],
typ = type function (table as table, optional options as record) as table
meta [Documentation.Name = "fxTableRemoveEmptyColumns", Documentation.Options = "function which deletes empty columns"]
,
result =  Value.ReplaceType(func, typ)
][result]
in
    Source
