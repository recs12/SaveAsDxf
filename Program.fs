(*
author: recs
date: 2021-02-05
summary: generate .dxf files from opened documents in application cad solidedge.
*)

open System
open System.IO
open SolidEdgeCommunity.Extensions
open SolidEdgeFramework



let saveflate sheet =
    let sheetMetalDocument = sheet :> SolidEdgePart.SheetMetalDocument
    let models = sheetMetalDocument.Models
    let flatPatternModels = sheetMetalDocument.FlatPatternModels

    //Checks if only one flatten exists in the cad.
    let hasFlate (flatPatternModels : SolidEdgePart.FlatPatternModels) : bool =
        let count : int = flatPatternModels.Count
        match count with
        | 0 -> false
        | 1 ->
            let flatPatternModel = flatPatternModels.Item(1)
            let flatPatternIsUpToDate = flatPatternModel.IsUpToDate
            match flatPatternIsUpToDate with
            | true -> true
            | false -> false
        | _ -> false

    let useFlatPattern : bool = hasFlate flatPatternModels

    let root = System.Environment.GetEnvironmentVariable("USERPROFILE")
    let destination = "\\Downloads\\solidedgeDxfs\\"
    let downloadsPath = root + destination
    let headDxf = System.IO.Path.ChangeExtension(sheetMetalDocument.Name, ".dxf")

    let saveAsDxf () =
        let sheetMetalDocument = sheet :> SolidEdgePart.SheetMetalDocument
        let outFile : string = System.IO.Path.Combine(downloadsPath, headDxf)
        printfn "%s, %b" sheetMetalDocument.Name useFlatPattern
        match useFlatPattern with
        | true ->
            if File.Exists(outFile) then ()
            else models.SaveAsFlatDXFEx(outFile, null, null, null, useFlatPattern)
        | false -> ()

    //Checks if the directory exist already before downloading the .dxf
    // if not it create a new folder.
    match Directory.Exists(downloadsPath) with
    | true
        -> saveAsDxf()
    | false ->
        Directory.CreateDirectory(downloadsPath) |> ignore
        saveAsDxf()
// One document to process
let onDxf app =
    let application = app :> SolidEdgeFramework.Application
    match application.ActiveDocumentType with
    | DocumentTypeConstants.igSheetMetalDocument ->
        let sh = application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false)
        saveflate sh |>ignore
    | _ -> printfn "This document is not sheet metal"


// Multiple documents to process
let multiDxf app =
    let application = app :> SolidEdgeFramework.Application
    let documents = application.Documents
    for doc in documents do
        let window = doc :?> SolidEdgeDocument
        window.Activate()
        match application.ActiveDocumentType with
        | DocumentTypeConstants.igSheetMetalDocument ->
            let sh = application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false)
            saveflate sh
        | _ -> printfn "%A this is not sheet metal" application.Name



[<STAThread>]
[<EntryPoint>]
    let main argv =

        try

            SolidEdgeCommunity.OleMessageFilter.Register()
            let application = SolidEdgeCommunity.SolidEdgeUtils.Connect(false)

            printfn "Would you like to get the .dxf of a sheet metal? (Press y/[Y] to proceed.):"
            printfn "(Note: key '*' for processing all opened sheet documents in batch)"
            let response = Console.ReadLine().ToLower()
            match response with
            | "y" ->
                printfn "Part-number, flatten"
                printfn "---"
                onDxf application
                printfn "..."

            | "*" ->
                printfn "Part-number, flatten"
                printfn "---"
                multiDxf application
                printfn "..."
            | _ -> printfn "Exit 0"

            printfn "Look in your Downloads folder for the dxfs."
            0

        finally
            SolidEdgeCommunity.OleMessageFilter.Unregister()
            printfn "Press any key to exit"
            Console.ReadKey() |> ignore
