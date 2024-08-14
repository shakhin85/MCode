(mylist as list)=>

let 
    mylist =  mylist,
    Conditions = List.Alternate(mylist, 1,1,1),
    Result = List.Alternate(mylist, 1,1,0),
    CombinedResult = {defaultVAlue}&Result,
    defaultVAlue = if Number.IsOdd(List.Count(mylist)) then List.Last(mylist) else null,
    returnIndex = List.PositionOf(Conditions, true)+1,
    returnItem = CombinedResult{returnIndex}
in 
    returnItem