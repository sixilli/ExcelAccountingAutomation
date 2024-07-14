// open System
// open Elmish
// open Avalonia
// open Avalonia.FuncUI.Hosts
// open Avalonia.FuncUI.Elmish
// open Avalonia.Controls.ApplicationLifetimes
// open Avalonia.Themes.Fluent
// open Yixin
//
// type MainWindow() as this =
//     inherit HostWindow()
//
//     do
//         base.Title <- "Tetris"
//         base.Width <- 450.0
//         base.Height <- 600.0
//
//         Program.mkProgram View.init View.update (View.view this)
//         |> Program.withHost this
//         |> Program.run
//
//
// type App() =
//     inherit Application()
//
//     override this.Initialize() =
//         this.Styles.Add(FluentTheme())
//         this.RequestedThemeVariant <- Styling.ThemeVariant.Dark
//
//     override this.OnFrameworkInitializationCompleted() =
//         match this.ApplicationLifetime with
//         | :? IClassicDesktopStyleApplicationLifetime as desktopLifetime -> desktopLifetime.MainWindow <- MainWindow()
//         | _ -> ()
//
//
// module Program =
//     [<EntryPoint>]
//     let main argv =
//         AppBuilder
//             .Configure<App>()
//             .UsePlatformDetect()
//             .UseSkia()
//             .StartWithClassicDesktopLifetime(argv)

module Program =
    open Yixin
    type Task =
    | GenerateReport
    | MoveInvoices
    | None
    type ProgramOptions = {
        inputDir : string
        outputDir: string
        invoiceDir: string
        task: Task
        file: string
    }
    let rec parseCommandLine args options =
        match args with
        | [] -> options
        | "--input"::tail | "-i"::tail ->
            match tail with
            | path::xs ->
                parseCommandLine xs { options with inputDir = path }
            | _ ->
                printfn "input requires a second parameter"
                exit -1
        | "--output"::tail | "-o"::tail ->
            match tail with
            | path::xs ->
                parseCommandLine xs { options with outputDir = path }
            | _ ->
                printfn "output requires a second parameter"
                exit -1
        | "--invoices"::tail | "-in"::tail ->
            match tail with
            | path::xs ->
                parseCommandLine xs { options with invoiceDir =  path }
            | _ ->
                printfn "invoice requires a second parameter"
                exit -1
        | task::tail ->
            match task.ToLower() with
            | "report" -> parseCommandLine tail { options with task = GenerateReport}
            | "move-invoices" -> parseCommandLine tail { options with task = MoveInvoices }
            | _ ->
                printfn $"unrecognized command: {task}, must be value of: report or move-invoices"
                exit -1
        
    [<EntryPoint>]
    let main argv =
        let args = System.Environment.GetCommandLineArgs() |> Array.skip 1 |> Array.toList
        if args.Length <= 0 then
            printfn "command line arguments cannot be empty, use report or move-invoices"
            exit -1

        let options = parseCommandLine args { inputDir=""; outputDir=""; invoiceDir = ""; task=None; file="" }
        
        let options =
            match (options.task, options.inputDir, options.outputDir) with
            | GenerateReport, inDir, outDir when inDir <> "" ->
                if outDir = "" then
                    { options with outputDir = inDir }
                else
                    options
            | MoveInvoices, inDir, outDir when inDir <> "" && outDir <> "" ->
                if options.invoiceDir = "" then
                    printfn "invoice directory is not set! use --invoices or -in"
                    exit -1
                else
                    options
            | _ ->
                printfn "input and output directories must be set! use -i and -o"
                exit -1
        
        match options.task with
        | GenerateReport ->
            try
                Processor.run options.inputDir options.outputDir
                printfn "process has finished running, press any key to exit"
                System.Console.ReadKey(true) |> ignore
                0
            with
            | ex ->
                printfn $"encountered an error: %s{ex.ToString()}"
                printfn "press any key to exit"
                System.Console.ReadKey(true) |> ignore
                ex.HResult
        | MoveInvoices ->
            try
                InvoicesDirectoryBuilder.run options.inputDir options.outputDir options.invoiceDir
                printfn "process has finished running, press any key to exit"
                System.Console.ReadKey(true) |> ignore
                0
            with
            | ex ->
                printfn $"encountered an error: %s{ex.ToString()}"
                printfn "press any key to exit"
                System.Console.ReadKey(true) |> ignore
                ex.HResult
        | _ ->
            printfn $"unrecognized command: {task}, must be value of: report or move-invoices"
            exit -1
