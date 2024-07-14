namespace Yixin


module InvoicesDirectoryBuilder =
    open System.IO
    open ClosedXML.Excel
    
    type Vendor = string
    type Invoice = string
    type VendorMap = Map<Vendor, Invoice seq>
    
    let walkRows (ws : IXLWorksheet) fn =
        let mutable it = ws.FirstRowUsed()
            
        while not (it.Cell(1).IsEmpty()) do
            fn it
            it <- it.RowBelow()
    let buildVendorMap ws =
        let mutable vendorMap : VendorMap = Map.empty
        
        let buildMap (row : IXLRow) =
            vendorMap <-
                let vendor = row.Cell(1).GetText()
                let invoice = row.Cell(2).GetText()
                match Map.tryFind vendor vendorMap with
                | Some invoices -> Map.add vendor (Seq.append invoices [invoice]) vendorMap
                | None -> Map.add vendor (seq [invoice]) vendorMap
        
        walkRows ws buildMap
        vendorMap
        
    type InvoicePath = {
        invoice : string
        path : string
    }
        
    type InvoiceSearchResults = {
        missingInvoices : VendorMap
        foundInvoices : Map<Vendor, InvoicePath seq>
    }
        
    // Return missing invoices!
    let searchForInvoices (vendorMap : VendorMap) (invoicePath : string) : InvoiceSearchResults =
        let mutable foundInvoices = Map.empty
        let mutable missingInvoices = Map.empty
        
        vendorMap
        |> Map.iter (fun vendor invoices ->
            let dirs = Directory.GetDirectories(invoicePath, $"*{vendor}*")
            
            match dirs.Length with
            | 0 ->
                  for invoice in invoices do
                      missingInvoices <-
                          match Map.tryFind vendor missingInvoices with
                          | Some invoices -> Map.add vendor (Seq.append invoices [invoice]) missingInvoices
                          | None -> Map.add vendor (seq [invoice]) missingInvoices
            | _ ->
                dirs
                |> Array.iter (fun dir ->
                    for invoice in invoices do
                        let fileMatches = Directory.GetFiles(dir, $"*{invoice}*", SearchOption.AllDirectories)
                        match fileMatches.Length with
                        | 0 ->
                            missingInvoices <-
                                match Map.tryFind vendor missingInvoices with
                                | Some invoices -> Map.add vendor (Seq.append invoices [invoice]) missingInvoices
                                | None -> Map.add vendor (seq [invoice]) missingInvoices
                        | _ ->
                            Array.iter (fun file ->
                                let invoicePath = { invoice = invoice; path = file }
                                foundInvoices <-
                                    match Map.tryFind vendor foundInvoices with
                                    | Some invoices -> Map.add vendor (Seq.append invoices [invoicePath]) foundInvoices
                                    | None -> Map.add vendor (seq [invoicePath]) foundInvoices
                            ) fileMatches
                    ) 
            ) 
        { missingInvoices = missingInvoices; foundInvoices = foundInvoices }
        
    
    // Create folders, even if they're missing!
    let buildVendorFoldersAndCopyInvoices (vendorData : VendorMap) (invoiceData : InvoiceSearchResults) outputPath =
        vendorData
        |> Map.iter(fun vendor _ ->
            let path = Path.Join(outputPath, vendor)
            let _dir = Directory.CreateDirectory(path)
            
            match Map.tryFind vendor invoiceData.foundInvoices with
            | Some invoices ->
                invoices
                |> Seq.iter (fun i ->
                    let copyPath = Path.Join(path, Path.GetFileName(i.path))
                    try
                        File.Copy(i.path, copyPath)
                    with
                    | :? IOException -> printfn $"file already exists: {copyPath}"
                    | ex ->
                        printfn $"encountered an unexpected exception, exiting: {ex.Message}"
                        exit -1
                )
            | None -> ()
        )
        
    let buildMissingInvoicesSheet outPath (invoiceData : InvoiceSearchResults) =
        let wb = new XLWorkbook()
        let ws = wb.Worksheets.Add("Missing Invoices")
        
        ws.Cell(1, 1).Value <- "Vendor"
        ws.Cell(1, 2).Value <- "Missing Invoice"
        let mutable row = 2
        invoiceData.missingInvoices
        |> Map.iter(fun vendor invoices ->
            invoices
            |> Seq.iter (fun i ->
                ws.Cell(row, 1).Value <- vendor
                ws.Cell(row, 2).Value <- i
                row <- row + 1
            )
        )
        
        ws.Cell(1, 4).Value <- "Vendor"
        ws.Cell(1, 5).Value <- "Found Invoice"
        ws.Cell(1, 6).Value <- "Path"
        let mutable row = 2
        invoiceData.foundInvoices
        |> Map.iter(fun vendor invoices ->
            invoices
            |> Seq.iter (fun i ->
                ws.Cell(row, 4).Value <- vendor
                ws.Cell(row, 5).Value <- i.invoice
                ws.Cell(row, 6).Value <- i.path
                row <- row + 1
            )
        )
        
        ws.Columns().AdjustToContents() |> ignore
        let savePath = Path.Join(outPath, "missing-invoices.xlsx")
        wb.SaveAs(savePath)
        

    let run (filePath : string) (outputPath : string) (invoicePath : string) =
        let inDir =
            if Directory.Exists(filePath) then
                filePath
            elif File.Exists(filePath) then
                Path.GetDirectoryName(filePath)
            else
                printfn $"invalid input path was given: {filePath}"
                exit -1
        
        let outDir =
            if Directory.Exists(outputPath) then
                outputPath
            elif File.Exists(filePath) then
                Path.GetDirectoryName(outputPath)
            elif (not (Directory.Exists(outputPath))) then
                Directory.CreateDirectory(outputPath).FullName
            else
                printfn $"invalid output path was given: {outputPath}"
                exit -1
                    
        let invoiceDir =
            if Directory.Exists(invoicePath) then
                invoicePath
            elif File.Exists(invoicePath) then
                Path.GetDirectoryName(invoicePath)
            else
                printfn $"invalid invoice path was given: {invoicePath}"
                exit -1
        
        let files = Directory.GetFiles(inDir, "*.xlsx")
        let selectedFile =
            match files.Length with
            | 0 ->
                printfn $"no valid files found in {inDir}"
                ""
            | _ ->
                printfn "Select a file to use to search for invoices"
                Seq.iteri (fun i file -> printfn $"{i+1}: {Path.GetFileName(file : string)}") files
                let mutable selectedFile = ""
                while selectedFile = "" do
                    try
                        printf "enter selection: "
                        let selection = System.Console.ReadLine().Trim() |> int
                        if selection < files.Length then
                            selectedFile <- files[selection-1]
                        else
                            printfn "invalid selection"
                        
                    with 
                    | :? System.FormatException -> 
                        printfn "invalid selection"
                selectedFile
                
        if selectedFile.Length <= 0 then
            printfn "no file selected, exiting"
            exit(-1)
            
        printfn $"searching for invoices in file: {selectedFile}"
        
        use workbook = new XLWorkbook(selectedFile)
        let vendorSheet = workbook.Worksheets.Worksheet(Constants.VendorInvoicesSheet)
        
        let vendorMap = buildVendorMap vendorSheet
        let searchResults = searchForInvoices vendorMap invoiceDir
        buildVendorFoldersAndCopyInvoices vendorMap searchResults outDir
        buildMissingInvoicesSheet outDir searchResults