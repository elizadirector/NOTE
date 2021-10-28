    ActiveWorkbook.Names.Add Name:="JPBuy", RefersToR1C1:= _
        "='20211022'!R4C2:R54C5"
    ActiveWorkbook.Names("JPBuy").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="JPSell", RefersToR1C1:= _
        "='20211022'!R4C6:R54C9"
    ActiveWorkbook.Names("JPSell").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="MLBuy", RefersToR1C1:= _
        "='20211022'!R4C10:R54C13"
    ActiveWorkbook.Names("MLBuy").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="MLSell", RefersToR1C1:= _
        "='20211022'!R4C14:R54C17"
    ActiveWorkbook.Names("MLSell").Comment = ""
    
    SELECT JPBuy.券商名稱 FROM JPBuy
WHERE 券商名稱
in (SELECT 券商名稱 FROM MLBuy)
