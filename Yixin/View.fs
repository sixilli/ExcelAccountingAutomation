namespace Yixin

module View =
    open Yixin.Components
    open Avalonia.FuncUI.DSL
    open Avalonia.FuncUI.Helpers
    open Avalonia.Controls
    open Avalonia.Layout
    open Avalonia.FuncUI.Hosts
    open Elmish

    type State =
        { title: string
          fileSelectorState: FileSelector.State }
        
    type Msg =
        | FileSelectorMsg of FileSelector.Msg
    let init () =
        let fsState, _ = FileSelector.init()
        { title = "hehe"
          fileSelectorState =  fsState }, Cmd.none
        
    let update (msg: Msg) (state: State) =
        match msg with
        | FileSelectorMsg fileSelectorMsg ->
            let s, cmd = FileSelector.update fileSelectorMsg state.fileSelectorState
            
            { state with fileSelectorState = s }, cmd

    let view (window: HostWindow) state dispatch =
        DockPanel.create [
              DockPanel.background "#222222"
              DockPanel.lastChildFill true
              DockPanel.children [
                    TextBlock.create [
                        TextBlock.dock Dock.Top
                        TextBlock.fontSize 48.0
                        TextBlock.verticalAlignment VerticalAlignment.Center
                        TextBlock.horizontalAlignment HorizontalAlignment.Center
                        TextBlock.text "hello :D"
                    ]
             ]
        ]
        |> generalize

