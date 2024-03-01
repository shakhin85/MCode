let
    Source = Excel.Workbook(File.Contents("V:\Findep\Incoming\test\DevOps\New_Mapping\Mapping_updated_on_12022024.xlsx"), null, true),
    CurrentMapping_Table = Source{[Item="CurrentMapping",Kind="Table"]}[Data],
    #"Removed Other Columns" = Table.SelectColumns(CurrentMapping_Table,{"Balance or PnL", "Rus Account", "Local Account Code", "Local Account Name_Rus", "Local Account Name_Eng", "Consolidation Account", "Consolidation Account Name", "FS Sub Line name", "FS code", "FS Line name", "Category", "Sign control"})
in
    #"Removed Other Columns"
