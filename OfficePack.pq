[
    readme="Небольшая библиотека для получения данных из файлов пакета MS Office, недоступных для штатных средств языка M. Как обычно всё открыто и бесплатно - пользуйтесь! @buchlotnik",

    fxUnzip=[func=(ZIP, optional options)=>
    [
    encod = [a=options[Encoding]?,b=if a = null then 866 else a][b],
    compr = (x)=> if x =0 then Compression.None else Compression.Deflate,

    u16 =BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
    u32 =BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),

    get=(x)=>BinaryFormat.Record(
        if Binary.Range(x,0,4)=#binary({0x50,0x4b,0x03,0x04})
        then [  Name=BinaryFormat.Text(u16(Binary.Range(x,26,2)),encod),
                Extr=BinaryFormat.Binary(u16(Binary.Range(x,28,2))),
                Value=BinaryFormat.Transform(BinaryFormat.Binary(u32(Binary.Range(x,18,4))),(y)=>Binary.Decompress(y,compr(u16(Binary.Range(x,8,2)))))]
        else []
    ),

    lst = BinaryFormat.List(BinaryFormat.Choice(BinaryFormat.Binary(30),get),(x)=>x<>[])(ZIP),
    to = Table.FromRecords(List.RemoveLastN(lst,1),type table [Name=text,Value=binary])][to],

typ =type function (ZIP as (type binary meta [Documentation.FieldCaption="ZIP - бинарное содержимое (zip, xlsx, docx, pptx и т.д.)"]), optional options as record) as table meta 
            [Documentation.Name = "fxUnzip (@buchlotnik)",
            Documentation.LongDescription ="функция считывает содержимое архива и возвращает таблицу с полями:<p> <b>Name</b> - имя файла (с путём к файлу) <p><b>Value</b> - бинарное содержимое файла. <p>Необязательный аргумент <b>options</b> на текущий момент поддерживает следующие поля: <p>1) <b>Encoding</b> - требуется для корректного чтения не латинских имён файлов (по умолчанию <b>866 - кириллица</b>)"],

result = Value.ReplaceType(func,typ)][result],

    fxUnzip2016=[ func = (ZIP, optional cp)=>
        [   ui16 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
            ui32 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            b = (x) => BinaryFormat.Binary(x),
            c = if cp is null then 866 else cp,
            f = (x) => try Binary.Decompress(x, Compression.Deflate) otherwise null,
            g = (x) => BinaryFormat.Transform(b(x),f),
            h = (x) => [head = BinaryFormat.Record([a = b(14), size=ui32, c=b(4), name=ui16, extra=ui16])(x),
                        body = BinaryFormat.Record( [ Name = BinaryFormat.Text(head[name],c),
                                                  Extras   = b(head[extra]),
                                                  Value  = g(head[size]),
                                                  flag  = true])][body],
            iftrue = BinaryFormat.Choice(b(26),h),
            iffalse = BinaryFormat.Record([flag = false]),
            getfile = BinaryFormat.Choice(ui32, (x)=> if x = 67324752 then iftrue else iffalse, type binary),
            filelist = BinaryFormat.List(getfile, each [flag] = true)(ZIP),
            remove = List.RemoveLastN(filelist,1),
            to = Table.FromRecords(remove, type table [Name = text, Value = binary])][to],
    typ = type function (
                ZIP as (type binary meta [Documentation.FieldCaption = "бинарное содержимое (zip,xlsx,docx,pptx и т.п.)"]),
                optional cp as (type number meta [Documentation.FieldCaption = "кодировка", Documentation.SampleValues = {866}])
                         )           
                as table meta [ 
                                Documentation.Name = "UnZip (@buchlotnik)",
                                Documentation.LongDescription = "Функция считывает содержимое zip-архива и возвращает таблицу с полями FileName - имя файла (с путём, если он в подпапке) и Content - бинарное содержимое конкретного файла. Кодировка требуется для корректного распознавания <b>не латинских</b> имён файлов (для офисных документов указывать необязательно), по умолчанию 866 - кодировка DOS <b>с поддержкой кириллицы</b>"        
                                ],
    result = Value.ReplaceType(func,typ)][result],

    fxExcelGetRowsAttributes= [func=(filebin,optional options)=>

   [sheetsonly = if options[SheetsOnly]?=true then true else false,
   fxUnzip = if options[Excel2016]?=true then fxUnzip2016 else fxUnzip,
      
      bin = Binary.Buffer(filebin),
    xml = List.Buffer(Table.SelectRows(fxUnzip(bin),(r)=>Text.StartsWith(r[Name],"xl/worksheets/sheet"))[Value]),
    wb = Table.Group(Excel.Workbook(bin,false),"Kind",{"tmp",(x)=>x}),
    xl = wb{[Kind="Sheet"]}[tmp],
    add = Table.AddIndexColumn(xl,"xml"),
    tr = Table.TransformColumns(add,{"xml",(x)=>xml{x}}),
    cmb = Table.CombineColumns(tr,{"xml","Data"},fxExcelSheetAddXmlInformation,"Data"),
    to = if sheetsonly then cmb else cmb & wb{[Kind="DefinedName"]}[tmp]][to],
    typ=type function(filebin as binary, optional options as record) as table meta 
    [Documentation.Name="fxExcelGetRowsAttributes (@buchlotnik)",
    Documentation.LongDescription="функция повторят действия функции Excel.Workbook, но к таблице каждого листа добавляет информацию об атрибутах строк"],
    result = Value.ReplaceType(func,typ)][result],

fxExcelSheetAddXmlInformation=[func=(x,optional options)=>
            [xml = Xml.Document(x{0}){0}[Value]{[Name="sheetData"]}[Value][Attributes],
            tr = List.Transform(xml,Record.FromTable),
            nms = List.Transform(xml,(x)=>x{[Name="r"]}?[Value]?),
            dict = Record.FromList(tr,nms),
            add=Table.AddIndexColumn(x{1},"Attributes",Number.From(nms{0})),
            tr1 = Table.TransformColumns(add,{"Attributes",(x)=>Record.FieldOrDefault(dict,Text.From(x))}),
            to = Table.ReorderColumns(tr1,{"Attributes"}&Table.ColumnNames(x{1}))][to],
    typ=type function(x as list,optional options as record) as table 
    meta [Documentation.Name="fxExcelSheetAddXmlInformation (@buchlotnik)",
        Documentation.LongDescription="вспомогательная функция, добавляющая к таблице листа дополнительную информацию из xml-разметки"],
    result=Value.ReplaceType(func,typ)][result],

fxTableRemoveEmptyColumns=[func =(table, optional options)=>
    [tbl = Table.Buffer(Table.Profile(table)),
    lst = Table.SelectRows(tbl,(r)=>r[Count]=r[NullCount])[Column],
    to = Table.RemoveColumns(table,lst)][to],
typ=type function (table as table, optional options as record) as table meta 
[Documentation.Name="fxTableRemoveEmptyColumns (@buchlotnik)",
Documentation.LongDescription="функция, удаляющая пустые столбцы из таблицы"],
result = Value.ReplaceType(func,typ)
][result],
fxGetMCode=(file)=>
    [u32 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
    f=(x)=>[a=Text.Split(x," = "),b={Text.Trim(Text.Replace(a{0},"shared","")),Text.Combine(List.Skip(a)," = ")}][b],
    from = fxUnzip(file){[Name="customXml/item1.xml"]}?[Value]?,
    bin = Binary.FromText(Xml.Document(from){0}[Value],BinaryEncoding.Base64),
    data = BinaryFormat.Choice(BinaryFormat.Binary(8),(x)=>BinaryFormat.Binary(u32(Binary.Range(x,4,4))))(bin),
    unz = fxUnzip(data){[Name="Formulas/Section1.m"]}[Value],
    txt = Text.FromBinary(unz)&"#(cr,lf)",
    splt = Text.Split(txt,";#(cr,lf)"),
    lst = List.Range(splt,1,List.Count(splt)-2),
    tbl = Table.FromList(lst,f,{"Name","Value"}),
    to = if from=null then #table({"Name","Value"},{{null,null}}) else tbl][to]

]
