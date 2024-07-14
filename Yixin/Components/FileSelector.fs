namespace Yixin.Components

module FileSelector =
    open Elmish
    open Avalonia.Controls
    open Avalonia.FuncUI.DSL
    open Avalonia.Platform.Storage
    
    type State =
        {
            count : int
            selectedFiles : string seq
        }
    let init() =
        {
            count = 0
            selectedFiles = [] 
        }
        , Cmd.none // no initial command
        
    type Msg =
        | Increment
        | Decrement
        | SetSelectedFiles of string seq
        
    let openFilePicker (window: Window) = async {
        let dialog = FilePickerOpenOptions()
        dialog.AllowMultiple <- false
        dialog.Title <- "Select a file"
        
        let! result = window.StorageProvider.OpenFilePickerAsync(dialog) |> Async.AwaitTask
        match result.Count with
        | 0 ->
            printfn "No file selected or multiple files selected"
            return (SetSelectedFiles [])
        | _ -> 
            let filePaths = Seq.map (fun (f : IStorageFile) -> f.TryGetLocalPath()) result
            return (SetSelectedFiles filePaths)
    }
    let update msg state =
        match msg with
        | Increment ->
            { state with
                count = state.count + 1
            }
            , Cmd.none
        | Decrement ->
            { state with
                count = state.count - 1
            }
            , Cmd.none
        | SetSelectedFiles filePaths -> { state with selectedFiles = filePaths }, Cmd.none

    let view (window : Window) (state: State) (dispatch: Msg -> unit)  =
        Button.create [
          Button.onClick(fun _ -> dispatch (openFilePicker(window) |> Async.RunSynchronously))
          Button.content("Select Files")
        ]
