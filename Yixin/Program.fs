open Yixin
[<EntryPoint>]
let main _ =
    try 
        Processor.run()
        printfn "process has finished running, press any key to exit"
        System.Console.ReadKey(true) |> ignore
        0
    with
    | ex ->
        printfn $"encountered an error: %s{ex.ToString()}"
        printfn "press any key to exit"
        System.Console.ReadKey(true) |> ignore
        ex.HResult
