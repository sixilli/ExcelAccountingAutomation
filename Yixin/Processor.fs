namespace Yixin

open System.IO
open System.Text.RegularExpressions
open ClosedXML.Excel

module Processor =
    type InvoiceData =
        { id: string
          amount: float
          vendor: string
          codeType: string
          amountLocations: IXLAddress list
          invoiceLocation: IXLAddress }
            
    type PhaseWrapper =
        { phaseLoc: IXLAddress
          invoiceData: InvoiceData list }
        
    [<Literal>]
    let ActualAmountColNum = 2
    [<Literal>]
    let InvoiceTotalColNum = 3
    [<Literal>]
    let InvoiceDiffColNum = 4
    [<Literal>]
    let FirstInvoiceColNum = 5
    // ending rows that are not needed
    let shouldSkip (phaseId : string) =
        let cleanVal = phaseId.ToLower()
        cleanVal = "total 10--total uses (all)" ||
        cleanVal = "no phase" ||
        cleanVal = "total all location" ||
        cleanVal.Contains("(all)")
        
    let walkRows (ws : IXLWorksheet) fn =
        let phaseLoc = 1
        
        let mutable it = ws.FirstRowUsed()
        // Skip first chunk of data as it's not useful
        while (it.Cell(phaseLoc).Value.GetText().Trim() <> "ALL LOCATION") do
            it <- it.RowBelow()
        it <- it.RowBelow()
            
        while not (it.Cell(phaseLoc).IsEmpty()) do
            let fullPhase = it.Cell(phaseLoc).Value.GetText().Trim()
            if not (shouldSkip fullPhase) then
                fn it
                
            it <- it.RowBelow()
        
    let walkPhases (ws : IXLWorksheet) fn =
        let phaseLoc = 1
        let amountLoc = 2
        let mutable it = ws.FirstRowUsed().RowUsed()
        // Skip first chunk of data as it's not useful
        while (it.Cell(phaseLoc).Value.GetText().Trim() <> "ALL LOCATION") do
            it <- it.RowBelow()
        it <- it.RowBelow()
            
        while not (it.Cell(phaseLoc).IsEmpty()) do
            let fullPhase = it.Cell(phaseLoc).Value.GetText().Trim()
            let code =
                match fullPhase.IndexOf("--") with
                | -1 -> fullPhase
                | i -> fullPhase.Substring(0, i) 
                
            let amountExists = it.Cell(amountLoc).Value.IsNumber
            
            if (not (shouldSkip fullPhase) && amountExists) then
                match code, it.Cell(amountLoc).Value.GetNumber() with
                | code, amount when amount > 0.0 ->
                    fn code (it.Cell(phaseLoc).Address)
                | _, _ -> ()
                
            it <- it.RowBelow()
            
    let collectNonZeroEntries (ws : IXLWorksheet) =
        let mutable map = Map.empty
        let workFn code phaseLoc =
            map <- map.Add(code, phaseLoc)
        walkPhases ws workFn
        map
        
    let private whiteSpaceRegex = Regex(@"\s+")
    let removeAllWhiteSpace input =
        whiteSpaceRegex.Replace(input, "")
        
        
    let combineInvoices (invoices: InvoiceData list) =
        invoices
        |> Seq.groupBy (_.id)
        |> Seq.map (fun (id, group) ->
            let firstInvoice = Seq.head group
            { id = id
              amount = group |> Seq.sumBy (_.amount)
              vendor = firstInvoice.vendor
              codeType = firstInvoice.codeType
              invoiceLocation = firstInvoice.invoiceLocation 
              amountLocations = 
                group 
                |> Seq.collect (_.amountLocations)
                |> Seq.toList }
        )
        |> List.ofSeq
        
    let findRelatedInvoices (glWs : IXLWorksheet) (map : Map<string, 'a>) =
        // util functions
        let getCellTextOrEmpty (cell : IXLCell) =
            if cell.IsEmpty() then "" else cell.Value.GetText()
            
        let getCellNumberOrZero (cell : IXLCell) =
            if cell.IsEmpty() then 0.0 else cell.Value.GetNumber()
        
        let table : IXLTable =
            let headersToFind = Set.ofSeq ["Doc"; "Phase"; "Credit"; "Debit"; "Vendor name"]
            let headerRow =
                glWs.RowsUsed()
                |> Seq.find (fun row ->
                    let foundHeaders = 
                        row.CellsUsed()
                        |> Seq.map (_.GetString())
                        |> Set.ofSeq
                    Set.isSubset headersToFind foundHeaders
                )
                
            // remove duplicate headers
            headerRow.CellsUsed()
            |> Seq.groupBy (_.Value.ToString())
            |> Seq.filter (fun (_, cells) -> Seq.length cells > 1)
            |> Seq.sortByDescending (fun (_, cells) -> Seq.last cells |> _.Address.ColumnNumber )
            |> Seq.iter (
                fun (_, cells) ->
                    Seq.last cells
                    |> fun cellToRemove ->
                       printfn $"Warning! found a duplicate '{cellToRemove.Value.ToString()}'column , will remove."
                       glWs.Column(cellToRemove.Address.ColumnNumber).Delete()
                )
            
            // find last cell
            let phaseCell = headerRow.Search("Phase") |> Seq.head
            let lastPhaseRow = glWs.Column(phaseCell.Address.ColumnNumber).LastCellUsed().Address.RowNumber
            let lastTableCell = glWs.Row(lastPhaseRow).LastCellUsed()
            
            // cache auto filter stuff
            glWs.AutoFilter.Clear() |> ignore
            
            glWs.Range(headerRow.FirstCellUsed().Address, lastTableCell.Address).CreateTable()
            
        let getColumnByName (columnName: string) =
            table.HeadersRow().Cells()
            |> Seq.find (fun cell -> cell.Value.GetText().Trim().ToLower() = columnName.ToLower())
        
        let docLoc = getColumnByName("Doc").Address.ColumnNumber
        let vendorLoc = getColumnByName("Vendor Name").Address.ColumnNumber
        let phaseIdLoc = getColumnByName("Phase").Address.ColumnNumber
        let creditLoc = getColumnByName("Credit").Address.ColumnNumber
        let debitLoc = getColumnByName("Debit").Address.ColumnNumber
            
        let addToNestedMap (phase: string) (vendor: string) (value: PhaseWrapper) (map: Map<string, Map<string, PhaseWrapper>>) =
            map
            |> Map.change phase (function
                | Some innerMap ->
                    if (innerMap.ContainsKey vendor) then
                        let currentWrapper = innerMap[vendor] 
                        let newWrapper = { currentWrapper with invoiceData = currentWrapper.invoiceData @ value.invoiceData  }
                        Some (innerMap |> Map.add vendor newWrapper)
                    else
                        Some (innerMap |> Map.add vendor value)
                | None -> 
                    Some (Map.empty |> Map.add vendor value)
            )
            
        // create map phase -> vendor -> invoice list
        let mutable outputMap: Map<string, Map<string, PhaseWrapper>> = Map.empty
        
        let processRow (row : IXLRangeRow) =
            let phaseCode = getCellTextOrEmpty(row.Cell(phaseIdLoc))
            
            if map.ContainsKey(phaseCode) && not (row.Cell(vendorLoc).IsEmpty()) then
                let phaseLoc = map[phaseCode]
                let debit = getCellNumberOrZero (row.Cell(debitLoc))
                let credit = getCellNumberOrZero (row.Cell(creditLoc))
                let vendor = getCellTextOrEmpty (row.Cell(vendorLoc))
                let invoiceId =
                    let d = getCellTextOrEmpty (row.Cell(docLoc))
                    let arr = d.Split(" ")
                    match arr with
                    | [| invoice |] -> invoice
                    | _ when arr.Length > 1 -> arr[arr.Length - 1]
                    | _ -> ":("
                let invoiceLoc = row.Cell(docLoc).Address
                
                let entry =
                    match credit, debit with
                    | c, d when c > d ->
                        {id = invoiceId; vendor = vendor; amount = debit + (credit * -1.0); invoiceLocation = invoiceLoc; amountLocations = [row.Cell(creditLoc).Address]; codeType = phaseCode }
                    | _ ->
                        {id = invoiceId; vendor = vendor; amount = debit + (credit * -1.0); invoiceLocation = invoiceLoc; amountLocations = [row.Cell(debitLoc).Address]; codeType = phaseCode }
                        
                outputMap <- addToNestedMap phaseCode vendor { phaseLoc = phaseLoc ; invoiceData = [entry]; } outputMap
            
        table.RowsUsed()
        |> Seq.skip 1
        |> Seq.iter processRow
        
        outputMap
       
        
    let updateReport (reportSheet : IXLWorksheet) (invoiceMap : Map<string, Map<string,PhaseWrapper>>) =
        let invoiceStartCol, invoiceRow, vendorRow = 5, 8, 9
        let mutable insertLoc = invoiceStartCol
        
        let createInvoiceFormula (address: IXLAddress) =
            $"='{address.Worksheet.Name}'!{address}"
        
        let formatCell (cell: IXLCell) =
            cell.Style.NumberFormat.Format <- "#,##0.00"
        
        let setInvoiceFormula (cell: IXLCell) (invoice: InvoiceData) =
            let sign = if invoice.amount < 0.0 then "-" else "+"
            cell.FormulaA1 <- $"{sign}'{invoice.amountLocations.Head.Worksheet.Name}'!{invoice.amountLocations.Head}"
        
        let processInvoice (phaseRow: int) (invoice: InvoiceData) =
            formatCell (reportSheet.Cell(phaseRow, insertLoc))
            setInvoiceFormula (reportSheet.Cell(phaseRow, insertLoc)) invoice
            reportSheet.Cell(vendorRow, insertLoc).Value <- invoice.vendor
            reportSheet.Cell(invoiceRow, insertLoc).FormulaA1 <- createInvoiceFormula invoice.invoiceLocation
            insertLoc <- insertLoc + 1
        
        let processMultipleInvoices (phaseRow: int) (invoices: seq<InvoiceData>) =
            formatCell (reportSheet.Cell(phaseRow, insertLoc))
            let formula = 
                invoices
                |> Seq.collect (fun invoice -> 
                    invoice.amountLocations 
                    |> Seq.map (fun loc -> 
                        let sign = if invoice.amount < 0.0 then "-" else "+"
                        $"{sign}'{loc.Worksheet.Name}'!{loc}"))
                |> String.concat ""
            reportSheet.Cell(phaseRow, insertLoc).FormulaA1 <- $"={formula.TrimStart('+')}"
            
            for invoice in invoices do
                reportSheet.Cell(vendorRow, insertLoc).Value <- invoice.vendor
                reportSheet.Cell(invoiceRow, insertLoc).FormulaA1 <- createInvoiceFormula invoice.invoiceLocation
                insertLoc <- insertLoc + 1
        
        let workFn phaseId _ =
            if invoiceMap.ContainsKey phaseId then
                invoiceMap[phaseId]
                |> Map.toSeq
                |> Seq.sortByDescending (fun (_, wrapper) -> wrapper.invoiceData.Length)
                |> Seq.iter (fun (_, wrapper) ->
                    let phaseRow = wrapper.phaseLoc.RowNumber
                    wrapper.invoiceData
                    |> Seq.groupBy (fun i -> i.id)
                    |> Seq.iter (fun (_, invoices) ->
                        match Seq.length invoices with
                        | 0 -> ()
                        | 1 -> processInvoice phaseRow (Seq.head invoices)
                        | _ -> processMultipleInvoices phaseRow invoices)) 
            
        walkPhases reportSheet workFn
        
        // merge cells in the same row with the same value
        let mutable itStart = reportSheet.Row(9).Cell(4).Address
        let mutable itEnd = reportSheet.Row(9).LastCellUsed().Address
        let endCopy = itEnd
        
        let mutable it = reportSheet.Range(itStart, itEnd).RangeUsed()
        let mutable currentValue = reportSheet.Row(9).Cell(4).Value
        let mutable currentValueCount = 1
        
        for cell in it.Cells() do
            match cell.Value with
            | c when c = currentValue ->
                currentValueCount <- currentValueCount + 1
                itEnd <- cell.Address
            | c when c <> currentValue && currentValueCount > 1 ->
                // update values for new value to check
                currentValue <- cell.Value
                currentValueCount <- 1
                // merge previous selection
                reportSheet.Range(itStart, itEnd).Merge() |> ignore
                // update new starting position since we have a new value
                itStart <- cell.Address
                currentValue <- cell.Value
                currentValueCount <- 1
            | c when c <> currentValue -> // no need to merge because we only had 1 before we got a new one
                itStart <- cell.Address
                currentValue <- cell.Value
                currentValueCount <- 1
            | _ when currentValueCount > 1 && cell.Address = endCopy -> 
                reportSheet.Range(itStart, itEnd).Merge() |> ignore
            | _ -> ()
                
    let searchForExcelFiles directory =
        Directory.GetFiles(directory, "*.xlsx")
        
    let applySumFormula ws =
        
        
        walkRows ws (fun row ->
            if row.Cell(ActualAmountColNum).Value.IsNumber then
                match row.Cell(ActualAmountColNum).Value.GetNumber() with
                | 0.0 -> row.Cell(InvoiceTotalColNum).Value <- 0.00
                | _ ->
                    let firstEntry = row.Cell(FirstInvoiceColNum).Address
                    let lastEntry = row.LastCellUsed().Address
                    let range = ws.Range(firstEntry, lastEntry)
                    
                    let calculatedTotal = Seq.fold (fun acc (cell : IXLCell) -> acc + cell.Value.GetNumber()) 0.0 (range.CellsUsed())
                    row.Cell(InvoiceTotalColNum).Value <- calculatedTotal
                    row.Cell(InvoiceTotalColNum).FormulaA1 <- $"=SUM({firstEntry}:{lastEntry})"
                    
                    if calculatedTotal <> 0.0 then
                        let actualTotalCell = row.Cell(ActualAmountColNum)
                        row.Cell(InvoiceDiffColNum).FormulaA1 <- $"{row.Cell(InvoiceTotalColNum).Address}-{actualTotalCell.Address}"
        )
        
    let run () =
        let dir = Directory.GetCurrentDirectory()
        let files = searchForExcelFiles dir
        match files.Length with
        | 0 -> printfn $"no valid files found in {dir}"
        | _ -> printfn $"attempting to do work on file: {files[0]}"
        
        use workbook = new XLWorkbook(files[0])
        let reportSheet = workbook.Worksheets.Worksheet(1)
        let glSheet = workbook.Worksheets.Worksheet(2)
        printfn "successfully loaded workbook"
        
        // processing pipeline
        let nonZeroMap = collectNonZeroEntries reportSheet
        let invoiceMap = findRelatedInvoices glSheet nonZeroMap
        updateReport reportSheet invoiceMap
        applySumFormula reportSheet
        
        let fileName = "sheet-builder-output.xlsx"
        let fileFull = Path.Join(dir, fileName)
        printfn $"finished processing, saving to: {fileFull}"
            
        workbook.SaveAs(fileFull) 